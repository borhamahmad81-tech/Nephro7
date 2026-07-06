"""
automation_engine.py
Drives Medvision Nephro (NEPHRO 7) via keyboard-shortcut automation.

IMPORTANT ARCHITECTURE NOTE: Nephro runs inside a Remote Desktop Connection
session. The local Windows OS can only see ONE real window - the RDP client
itself ("NEPHRO 7 (Remote)") - it has zero visibility into windows that
exist inside the remote session (the search dialog, note editor, verification
prompt all live on the remote server, not locally). So:

  - Focus verification (are we even looking at the right RDP window) uses
    real Windows APIs (win32gui) - this works fine, the RDP client is a
    normal local window.
  - Detecting state CHANGES INSIDE the session (dialog opened/closed) CANNOT
    use window titles - it uses image template matching on screenshots
    instead, since RDP is just pixels either way and screenshots see
    whatever is rendered, local or remote.

Required template images (in the same folder as this script):
  search_patients_dialog.png   - crop of the "Search patients" title bar
  progress_notes_editor.png    - crop of the "PROGRESS NOTES" editor title bar
  verification_dialog.png      - crop of the "Verification..." dialog title bar
  doctors_progress_tab.png     - optional, only if the tab isn't default-active
"""

import time
import threading
import pyautogui
import pyperclip

pyautogui.FAILSAFE = True

TEMPLATE_SEARCH_PATIENTS = "search_patients_dialog.png"
TEMPLATE_PROGRESS_NOTES = "progress_notes_editor.png"
TEMPLATE_VERIFICATION = "verification_dialog.png"
TEMPLATE_PROGRESS_TAB = "doctors_progress_tab.png"  # optional fallback only

IMAGE_ASSETS = [
    TEMPLATE_SEARCH_PATIENTS,
    TEMPLATE_PROGRESS_NOTES,
    TEMPLATE_VERIFICATION,
    TEMPLATE_PROGRESS_TAB,
]


class AbortRequested(Exception):
    pass


class AutomationEngine:
    def __init__(self, log_callback, image_dir=".", confidence=0.8):
        self.log = log_callback
        self.image_dir = image_dir
        self.confidence = confidence

        self._abort_flag = threading.Event()
        self._pause_flag = threading.Event()

        # Poll timeouts (ceilings, not fixed waits) - overridable from GUI
        self.search_open_timeout = 15
        self.patient_load_timeout = 90     # covers cloud/RDP lag opening the record
        self.note_editor_timeout = 30
        self.verify_open_timeout = 20
        self.save_commit_timeout = 90      # covers cloud/RDP lag on save/commit
        self.poll_interval = 0.3

        self.assume_progress_tab_active = True
        self.guided_mode = True
        self.password = ""
        self.on_ready_to_sign = None

        # pygetwindow/win32gui used ONLY to verify the local RDP client window
        # has real OS focus - not used for anything happening inside the
        # remote session (see module docstring).
        self._gw = None
        try:
            import pygetwindow as gw
            self._gw = gw
        except ImportError:
            pass

    # ---------------- control ----------------

    def request_abort(self):
        self._abort_flag.set()
        self._pause_flag.clear()

    def _check_abort(self):
        if self._abort_flag.is_set():
            raise AbortRequested()

    def _asset(self, filename):
        import os
        return os.path.join(self.image_dir, filename)

    # ---------------- local RDP-client focus verification ----------------

    def _focus_nephro_window(self):
        """
        Brings the local RDP client window (titled containing "NEPHRO") to
        the foreground and VERIFIES it actually got focus via GetForegroundWindow
        before returning. Raises RuntimeError if it can't confirm - stopping
        rather than sending keystrokes into the void.
        """
        try:
            import win32gui
            import win32con
            import ctypes
        except ImportError:
            raise RuntimeError(
                "pywin32 is required to reliably focus the Nephro window "
                "(pip install pywin32) - cannot safely proceed without it."
            )

        hwnd = None

        def _enum(h, _):
            nonlocal hwnd
            if hwnd is None and "NEPHRO" in win32gui.GetWindowText(h).upper():
                hwnd = h

        win32gui.EnumWindows(_enum, None)

        if not hwnd:
            raise RuntimeError(
                "Could not find a window with 'NEPHRO' in its title at all - "
                "is the Remote Desktop session to Nephro open and connected?"
            )

        for _ in range(3):
            self._check_abort()

            # If it's already correctly focused, don't touch anything -
            # sending an unneeded Alt keypress can accidentally activate
            # the remote app's own menu bar, which then swallows F3/etc.
            if win32gui.GetForegroundWindow() == hwnd:
                self._settle_rdp_input_focus(hwnd)
                return

            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

            ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)  # Alt down
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
            ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)  # Alt up
            time.sleep(0.3)

            if win32gui.GetForegroundWindow() == hwnd:
                self._settle_rdp_input_focus(hwnd)
                return

            time.sleep(0.3)

        raise RuntimeError(
            "Could not confirm the Nephro/RDP window actually received focus "
            "after 3 attempts - aborting rather than guessing. Try clicking "
            "on the Nephro RDP window manually right before pressing Start."
        )

    def _settle_rdp_input_focus(self, hwnd):
        """
        OS-level foreground focus doesn't always guarantee the remote RDP
        session is actually capturing keyboard input yet. A real click
        inside the window content area forces that, and a safety Escape
        clears any menu accidentally activated by the Alt-key focus trick.
        """
        import win32gui

        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            cx = (left + right) // 2
            cy = top + 12  # title bar strip, not the window content - avoids
                            # accidentally clicking a patient row/button inside
            pyautogui.click(cx, cy)
            time.sleep(0.2)
            pyautogui.press("escape")  # clear any accidentally-activated menu
            time.sleep(0.2)
        except Exception as exc:
            self.log("warn", f"Could not settle RDP input focus: {exc}")

    # ---------------- image-based detection (inside the RDP session) ----------------

    def _find_on_screen(self, template_filename):
        path = self._asset(template_filename)
        try:
            return pyautogui.locateOnScreen(path, confidence=self.confidence)
        except pyautogui.ImageNotFoundException:
            return None
        except Exception:
            return None

    def _wait_for_template(self, template_filename, should_appear, timeout, label=""):
        """
        Poll the screen (via screenshot + template match) until the given
        template image is found (should_appear=True) or disappears
        (should_appear=False), or timeout - in which case this RAISES,
        aborting the run rather than silently continuing.
        """
        deadline = time.time() + timeout
        while time.time() < deadline:
            self._check_abort()
            found = self._find_on_screen(template_filename) is not None
            if found == should_appear:
                return
            time.sleep(self.poll_interval)

        verb = "appear" if should_appear else "disappear"
        raise RuntimeError(
            f"'{template_filename}' never seemed to {verb}"
            + (f" ({label})" if label else "")
            + f" within {timeout}s - stopping rather than guessing the app is done. "
            "If this keeps happening even though the step clearly worked on screen, "
            "the template image likely doesn't match your actual resolution/zoom - "
            "recapture it directly from your live screen."
        )

    def _ensure_progress_tab(self):
        if self.assume_progress_tab_active:
            return
        box = self._find_on_screen(TEMPLATE_PROGRESS_TAB)
        if box:
            pyautogui.click(pyautogui.center(box))
            time.sleep(0.5)
        else:
            self.log("warn", "Progress Notes tab image not found - assuming it's already active.")

    def check_assets(self):
        import os
        return {name: os.path.isfile(self._asset(name)) for name in IMAGE_ASSETS}

    # ---------------- per-patient pipeline ----------------

    def process_patient(self, record):
        self._check_abort()
        self.log("info", f"--- Patient {record.patient_id} ---")

        self._focus_nephro_window()

        # 1. Open patient search (F3)
        pyautogui.press("f3")
        self._wait_for_template(
            TEMPLATE_SEARCH_PATIENTS, True, self.search_open_timeout,
            label="search dialog open"
        )
        self._check_abort()

        # 2. Type ID, Enter to search, Enter to confirm/open result
        pyautogui.typewrite(record.patient_id, interval=0.03)
        pyautogui.press("enter")
        time.sleep(1.0)  # brief settle for the in-list search results
        pyautogui.press("enter")
        self._check_abort()

        # 3. Wait for the search dialog to disappear - this IS the
        #    patient-loaded signal, adaptive to RDP/cloud lag.
        self._wait_for_template(
            TEMPLATE_SEARCH_PATIENTS, False, self.patient_load_timeout,
            label="patient record opened",
        )
        self._check_abort()

        # 4. Ensure the right tab is active (usually already default)
        self._ensure_progress_tab()
        self._check_abort()

        # 5. New Progress Note entry
        pyautogui.hotkey("ctrl", "insert")
        self._wait_for_template(
            TEMPLATE_PROGRESS_NOTES, True, self.note_editor_timeout,
            label="note editor opened"
        )
        self._check_abort()

        # 6. Keep header line, clear the template body below it
        pyautogui.hotkey("ctrl", "home")
        pyautogui.press("down")
        pyautogui.hotkey("ctrl", "shift", "end")
        pyautogui.press("delete")
        pyautogui.press("enter")
        pyautogui.press("enter")
        self._check_abort()

        # 7. Paste the clinical note
        pyperclip.copy(record.note_text)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.8)
        self._check_abort()

        # 8. Open sign-off dialog
        pyautogui.hotkey("ctrl", "alt", "s")
        self._wait_for_template(
            TEMPLATE_VERIFICATION, True, self.verify_open_timeout,
            label="verification dialog opened"
        )
        self._check_abort()

        if self.guided_mode:
            self.log("warn", f"Paused before signing patient {record.patient_id}.")
            if self.on_ready_to_sign:
                self.on_ready_to_sign(record)
            self._wait_while_paused()
            self._check_abort()

        # 9. Type password, Enter to commit
        pyautogui.typewrite(self.password, interval=0.02)
        pyautogui.press("enter")

        # 10. Confirm the save actually committed (dialog disappears)
        self._wait_for_template(
            TEMPLATE_VERIFICATION, False, self.save_commit_timeout,
            label="save committed",
        )

        self.log("success", f"Patient {record.patient_id} saved.")

    def _wait_while_paused(self):
        self._pause_flag.set()
        while self._pause_flag.is_set():
            self._check_abort()
            time.sleep(0.2)

    def resume(self):
        self._pause_flag.clear()

    # ---------------- full run ----------------

    def run(self, records, on_done):
        self._abort_flag.clear()
        total = len(records)
        completed = 0
        try:
            for record in records:
                self.process_patient(record)
                completed += 1
            self.log("success", f"Run complete: {completed}/{total} patients processed.")
        except AbortRequested:
            self.log("warn", f"Aborted by user after {completed}/{total} patients.")
        except Exception as exc:
            self.log("error", f"Stopped due to error after {completed}/{total}: {exc}")
        finally:
            on_done(completed, total)
