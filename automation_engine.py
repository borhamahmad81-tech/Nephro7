"""
automation_engine.py
Drives Medvision Nephro (NEPHRO 7) via keyboard-shortcut automation.

Instead of fixed sleep delays (unreliable over a cloud/remote connection),
each step polls for a real window-state change using distinct window
titles observed in the app:

  F3                    -> "Search patients" window appears
  <ID> + Enter + Enter   -> "Search patients" window closes (patient opened)
  Ctrl+Insert            -> "PROGRESS NOTES" editor window appears
  Ctrl+Alt+S             -> "Verification..." window appears
  <password> + Enter     -> "Verification..." window closes (save committed)

Polling has a generous timeout ceiling per step; if a window never reaches
the expected state within that ceiling, the step logs a warning and falls
back to a short fixed delay rather than hanging forever.
"""

import time
import threading
import pyautogui
import pyperclip

pyautogui.FAILSAFE = True

IMAGE_ASSETS = [
    "doctors_progress_tab.png",  # optional fallback only
]

# Window titles observed in the real app - adjust here if your build differs.
WIN_SEARCH_PATIENTS = "Search patients"
WIN_PROGRESS_NOTES = "PROGRESS NOTES"
WIN_VERIFICATION = "Verification"


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
        self.patient_load_timeout = 90     # covers cloud lag opening the record
        self.note_editor_timeout = 30
        self.verify_open_timeout = 20
        self.save_commit_timeout = 90      # covers cloud lag on save/commit
        self.poll_interval = 0.3

        self.assume_progress_tab_active = True
        self.guided_mode = True
        self.password = ""
        self.on_ready_to_sign = None

        self._gw = None
        try:
            import pygetwindow as gw
            self._gw = gw
        except ImportError:
            self.log("warn", "pygetwindow not installed - falling back to fixed delays.")

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

    # ---------------- window polling ----------------

    def _window_exists(self, title_substring):
        if not self._gw:
            return None  # unknown - caller should fall back
        try:
            return bool(self._gw.getWindowsWithTitle(title_substring))
        except Exception:
            return None

    def _wait_for_window(self, title_substring, should_exist, timeout, label=""):
        """
        Poll until a window containing `title_substring` matches the
        should_exist state, or timeout - in which case this RAISES, aborting
        the run. Silently continuing after a failed detection is exactly
        what caused false "saved" reports with nothing actually done, so
        this now fails loudly instead.
        """
        if self._gw is None:
            raise RuntimeError(
                "pygetwindow is required to verify each step actually happened "
                "(pip install pygetwindow) - cannot safely proceed without it."
            )

        deadline = time.time() + timeout
        while time.time() < deadline:
            self._check_abort()
            state = self._window_exists(title_substring)
            if state == should_exist:
                return

            time.sleep(self.poll_interval)

        verb = "appear" if should_exist else "close"
        raise RuntimeError(
            f"'{title_substring}' window never seemed to {verb}"
            + (f" ({label})" if label else "")
            + f" within {timeout}s - stopping rather than guessing the app is done."
        )

    def _focus_nephro_window(self):
        """
        Brings the NEPHRO window to the foreground and VERIFIES it actually
        got focus before returning. Windows blocks background processes from
        silently stealing foreground focus, so a plain .activate() call can
        fail silently - this uses the standard Alt-key workaround, then
        checks GetForegroundWindow() to confirm. Raises RuntimeError (which
        aborts the run) if it cannot confirm focus, rather than continuing
        to send keystrokes into the void.
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
        if self._gw:
            windows = self._gw.getWindowsWithTitle("NEPHRO")
            if windows:
                hwnd = windows[0]._hWnd
        if hwnd is None:
            hwnd = win32gui.FindWindow(None, None)  # placeholder, real search below
            hwnd = None
            def _enum(h, _):
                nonlocal hwnd
                if "NEPHRO" in win32gui.GetWindowText(h).upper():
                    hwnd = h
            win32gui.EnumWindows(_enum, None)

        if not hwnd:
            raise RuntimeError(
                "Could not find the NEPHRO window at all - is Medvision Nephro open?"
            )

        for attempt in range(3):
            self._check_abort()
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

            # Alt-key workaround for Windows' foreground-lock restriction
            ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)  # Alt down
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
            ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)  # Alt up
            time.sleep(0.3)

            if win32gui.GetForegroundWindow() == hwnd:
                return  # confirmed - safe to send keystrokes now

            time.sleep(0.3)

        raise RuntimeError(
            "Could not confirm Nephro actually received focus after 3 attempts - "
            "aborting rather than sending keystrokes into the void. "
            "Try clicking on the Nephro window manually right before pressing Start."
        )

    # ---------------- image fallback (tab only) ----------------

    def _locate(self, filename, timeout=4.0, interval=0.4):
        deadline = time.time() + timeout
        path = self._asset(filename)
        while time.time() < deadline:
            self._check_abort()
            try:
                box = pyautogui.locateOnScreen(path, confidence=self.confidence)
            except pyautogui.ImageNotFoundException:
                box = None
            if box:
                return pyautogui.center(box)
            time.sleep(interval)
        return None

    def _ensure_progress_tab(self):
        if self.assume_progress_tab_active:
            return
        point = self._locate("doctors_progress_tab.png", timeout=4.0)
        if point:
            pyautogui.click(point)
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
        self._wait_for_window(
            WIN_SEARCH_PATIENTS, True, self.search_open_timeout,
            label="search dialog open"
        )
        self._check_abort()

        # 2. Type ID, Enter to search, Enter to confirm/open result
        pyautogui.typewrite(record.patient_id, interval=0.03)
        pyautogui.press("enter")
        time.sleep(1.0)  # brief settle for the in-list search results (not network bound)
        pyautogui.press("enter")
        self._check_abort()

        # 3. Wait for the search dialog to close - this IS the patient-loaded signal,
        #    adaptive to however slow the cloud connection is today.
        self._wait_for_window(
            WIN_SEARCH_PATIENTS, False, self.patient_load_timeout,
            label="patient record opened",
        )
        self._check_abort()

        # 4. Ensure the right tab is active (usually already default)
        self._ensure_progress_tab()
        self._check_abort()

        # 5. New Progress Note entry
        pyautogui.hotkey("ctrl", "insert")
        self._wait_for_window(
            WIN_PROGRESS_NOTES, True, self.note_editor_timeout,
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
        self._wait_for_window(
            WIN_VERIFICATION, True, self.verify_open_timeout,
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

        # 10. Confirm the save actually committed (dialog closes) - adaptive
        #     to cloud lag instead of a blind "save delay".
        self._wait_for_window(
            WIN_VERIFICATION, False, self.save_commit_timeout,
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
