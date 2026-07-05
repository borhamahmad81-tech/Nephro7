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

    def _wait_for_window(self, title_substring, should_exist, timeout, fallback_delay=2.0, label=""):
        """
        Poll until a window containing `title_substring` matches the
        should_exist state, or timeout. Returns True if the state was
        confirmed, False if it timed out (caller should treat as a warning,
        not necessarily a hard failure).
        If pygetwindow isn't available at all, just sleeps fallback_delay.
        """
        if self._gw is None:
            time.sleep(fallback_delay)
            return False

        deadline = time.time() + timeout
        while time.time() < deadline:
            self._check_abort()
            state = self._window_exists(title_substring)
            if state == should_exist:
                return True
            time.sleep(self.poll_interval)

        verb = "appear" if should_exist else "close"
        self.log(
            "warn",
            f"Timed out waiting for '{title_substring}' window to {verb}"
            + (f" ({label})" if label else "")
            + f" after {timeout}s - continuing anyway.",
        )
        time.sleep(fallback_delay)
        return False

    def _focus_nephro_window(self):
        if not self._gw:
            return
        try:
            windows = self._gw.getWindowsWithTitle("NEPHRO")
            if not windows:
                self.log("warn", "Could not find a NEPHRO window to focus - is it open?")
                return
            win = windows[0]
            if win.isMinimized:
                win.restore()
            win.activate()
            time.sleep(0.3)
        except Exception as exc:
            self.log("warn", f"Window focus attempt failed: {exc}")

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
            fallback_delay=2.0, label="search dialog open"
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
            fallback_delay=self.patient_load_timeout if self._gw is None else 2.0,
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
            fallback_delay=2.0, label="note editor opened"
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
            fallback_delay=1.5, label="verification dialog opened"
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
            fallback_delay=2.0, label="save committed",
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
