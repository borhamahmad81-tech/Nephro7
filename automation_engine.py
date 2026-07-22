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
  new_note_icon.png            - crop of the new-note (blank page) icon TOGETHER
                                 with part of the "PROGRESS NOTES" panel title,
                                 so it matches ONLY inside the Progress Notes
                                 panel and NOT the identical icon in Doctors
                                 Orders. This is clicked to open a new note
                                 instead of pressing Ctrl+Insert, because the
                                 keystroke depends on which pane has focus
                                 (it also fires in Doctors Orders) whereas a
                                 located icon click is focus-independent.
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
TEMPLATE_NEW_NOTE_ICON = "new_note_icon.png"  # icon + panel-title anchor
TEMPLATE_PROGRESS_TAB = "doctors_progress_tab.png"  # optional fallback only

# Required templates that MUST exist for a run to be safe.
REQUIRED_ASSETS = [
    TEMPLATE_SEARCH_PATIENTS,
    TEMPLATE_PROGRESS_NOTES,
    TEMPLATE_VERIFICATION,
    TEMPLATE_NEW_NOTE_ICON,
]

IMAGE_ASSETS = REQUIRED_ASSETS + [TEMPLATE_PROGRESS_TAB]


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
        self.search_open_timeout = 30
        self.patient_load_timeout = 90     # covers cloud/RDP lag opening the record
        self.note_editor_timeout = 90
        self.verify_open_timeout = 30
        self.save_commit_timeout = 90      # covers cloud/RDP lag on save/commit
        self.poll_interval = 0.3

        self.assume_progress_tab_active = True
        self.guided_mode = True
        self.password = ""
        self.on_ready_to_sign = None
        self.post_load_settle = 2.0  # extra buffer after patient search closes,
                                      # before doing anything patient-specific -
                                      # combats stale-previous-patient race condition

        # Where within the matched new-note-icon template to click, as
        # fractions of the matched box (0..1). The template includes the icon
        # plus some panel title/tab, so the geometric centre may land on text
        # instead of the icon. THIS crop is "PROGRESS NOTES" text on top with
        # the icon in the BOTTOM portion, so we click low in the box. If the
        # click lands slightly off the icon, nudge these: smaller y = higher,
        # larger y = lower.
        self.icon_click_x_frac = 0.50
        self.icon_click_y_frac = 0.78

        # Sentinel word used to confirm the template body was actually deleted
        # before pasting. It must be a word that appears at/near the BOTTOM of
        # the template (so if it's gone, the whole body from top to bottom is
        # gone) and does NOT appear in the one-line patient header. The user's
        # template runs "from Complaint ... to Educational", so 'Educational'
        # (the last template line) is the default. Checked BEFORE paste, so a
        # note that happens to contain the word later cannot cause a false
        # positive.
        self.clear_sentinel = "Educational"

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

    def _park_cursor_safe(self):
        """
        Move the cursor away from the screen corners before doing anything.

        PyAutoGUI's FAILSAFE treats "mouse in a corner" as a panic/abort
        signal and raises on the NEXT pyautogui action. If the user leaves
        the cursor parked in a corner when a run starts, the very first
        click aborts instantly (this is exactly what bit us in the field).

        We use win32api.SetCursorPos (NOT pyautogui.moveTo) so this move
        itself does not go through the fail-safe check and cannot abort.
        Parking roughly centre-screen keeps the real fail-safe intact as a
        manual kill switch (user slams mouse into a corner to stop), while
        preventing accidental self-triggering at the start of each action.
        """
        try:
            import win32api
            w, h = pyautogui.size()
            win32api.SetCursorPos((w // 2, h // 2))
        except Exception as exc:
            # Non-fatal: if this fails we simply proceed and let the normal
            # fail-safe behaviour stand.
            self.log("warn", f"Could not park cursor centre-screen: {exc}")

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
            # Get the cursor out of any fail-safe corner FIRST, otherwise the
            # click below can abort the whole run on its fail-safe check.
            self._park_cursor_safe()
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
        except Exception as exc:
            # Previously this was silently swallowed, which made a broken
            # OpenCV/matching setup look identical to "template doesn't
            # match" - always log the real reason so that distinction is
            # visible instead of guessing at templates forever.
            self.log("error", f"Image match error on '{template_filename}': {type(exc).__name__}: {exc}")
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
            box = self._find_on_screen(template_filename)
            found = box is not None
            if found == should_appear:
                return box if should_appear else None
            time.sleep(self.poll_interval)

        verb = "appear" if should_appear else "disappear"
        shot_path = self._save_timeout_screenshot(label or template_filename)
        raise RuntimeError(
            f"'{template_filename}' never seemed to {verb}"
            + (f" ({label})" if label else "")
            + f" within {timeout}s - stopping rather than guessing the app is done. "
            + (f"Saved a screenshot of what was on screen to: {shot_path}. " if shot_path else "")
            + "If this keeps happening even though the step clearly worked on screen, "
            "the template image likely doesn't match your actual resolution/zoom - "
            "recapture it directly from your live screen."
        )

    def _save_timeout_screenshot(self, label):
        """
        Dump a full-screen screenshot to disk whenever a detection step times
        out, so a failed run comes back with a picture of what was actually
        on screen instead of a guess. Returns the path, or None on failure.
        """
        import os
        try:
            safe = "".join(c if c.isalnum() else "_" for c in label)[:40]
            stamp = time.strftime("%Y%m%d_%H%M%S")
            path = os.path.join(self.image_dir, f"timeout_{safe}_{stamp}.png")
            pyautogui.screenshot().save(path)
            self.log("warn", f"Saved timeout screenshot: {path}")
            return path
        except Exception as exc:
            self.log("warn", f"Could not save timeout screenshot: {exc}")
            return None

    def _ensure_progress_tab(self):
        if self.assume_progress_tab_active:
            return
        box = self._find_on_screen(TEMPLATE_PROGRESS_TAB)
        if box:
            pyautogui.click(pyautogui.center(box))
            time.sleep(0.5)
        else:
            self.log("warn", "Progress Notes tab image not found - assuming it's already active.")

    def _click_new_note_icon(self, timeout):
        """
        Open a new progress note by CLICKING the new-note icon, rather than
        pressing Ctrl+Insert.

        Why not Ctrl+Insert: that keystroke goes to whichever pane currently
        has keyboard focus. Over RDP, right after a focus switch, that pane
        is not reliably the Progress Notes panel - the identical shortcut
        also creates a new entry in Doctors Orders if Orders has focus. So
        the keystroke "works sometimes" depending on the last click. Clicking
        a located on-screen icon does not depend on focus at all.

        Why the template is icon + panel-title: the same blank-page icon
        appears in BOTH the Progress Notes panel and Doctors Orders. Matching
        the icon ALONE could click the Orders one (locateOnScreen returns the
        first match in scan order). The template therefore includes part of
        the unique "PROGRESS NOTES" panel title next to the icon, so it only
        matches inside the Progress Notes panel.

        This also doubles as an OBSTRUCTION GUARD: if the anchored icon is not
        visible, either the Progress Notes panel is covered by another window
        or Nephro is not really on top. In that case we STOP rather than click
        blind - a blind click here could land anywhere.
        """
        box = self._wait_for_template(
            TEMPLATE_NEW_NOTE_ICON, True, timeout,
            label="progress-notes new-note icon (anchor visible)"
        )
        self._park_cursor_safe()
        # The template crop is the new-note icon together with part of the
        # PROGRESS NOTES panel title/tab (used so it only matches inside the
        # Progress Notes panel, never the identical icon in Doctors Orders).
        # The ICON is in the crop; the title/tab text is the rest. We must
        # click the ICON, not the text. Which part of the box the icon sits in
        # depends on how the crop was made, so both offsets are tunable:
        #   icon_click_x_frac / icon_click_y_frac are fractions (0..1) of the
        #   matched box width/height giving the click point.
        # Default targets the upper-left region, where the icon sits when the
        # crop is icon-above-title (tall, narrow box) or icon-left-of-title.
        bx, by = int(box.left), int(box.top)
        bw, bh = int(box.width), int(box.height)
        click_x = bx + max(6, int(bw * self.icon_click_x_frac))
        click_y = by + max(6, int(bh * self.icon_click_y_frac))
        self.log("info", f"Clicking new-note icon at ({click_x}, {click_y}) "
                         f"within match Box(left={bx}, top={by}, w={bw}, h={bh}).")
        pyautogui.click(click_x, click_y)
        time.sleep(0.4)

    def check_assets(self):
        import os
        return {name: os.path.isfile(self._asset(name)) for name in IMAGE_ASSETS}

    def check_required_assets(self):
        import os
        return {name: os.path.isfile(self._asset(name)) for name in REQUIRED_ASSETS}

    # ---------------- OCR-based patient ID verification ----------------

    @staticmethod
    def _configure_tesseract():
        """
        Points pytesseract at the bundled Tesseract binary when running as a
        frozen EXE (PyInstaller --add-binary/--add-data), so no separate
        install is needed on the clinic PC. Falls back to whatever's on
        PATH when running from source.
        """
        import sys
        import os
        import pytesseract

        if getattr(sys, "frozen", False):
            base = sys._MEIPASS
            exe_path = os.path.join(base, "tesseract", "tesseract.exe")
            tessdata_path = os.path.join(base, "tesseract", "tessdata")
            if os.path.isfile(exe_path):
                pytesseract.pytesseract.tesseract_cmd = exe_path
                os.environ["TESSDATA_PREFIX"] = tessdata_path

    def _read_region_text(self, left, top, width, height):
        import pytesseract
        self._configure_tesseract()
        region = (int(left), int(top), int(width), int(height))
        shot = pyautogui.screenshot(region=region)
        return pytesseract.image_to_string(shot)

    def _wait_for_correct_patient_header(self, expected_id, editor_box, timeout=None):
        """
        Reads the note's header line (via OCR) and waits until it actually
        contains the expected patient ID, instead of guessing a fixed delay.
        This is the real fix for the "previous patient's note opened by
        mistake" race condition - it keeps checking for as long as needed
        (bounded by timeout) rather than assuming any particular load time.
        Raises RuntimeError (aborting the run) rather than risk pasting
        into the wrong patient's chart.
        """
        if timeout is None:
            timeout = self.patient_load_timeout

        # editor_box coordinates come back from pyautogui as numpy int64, which
        # pyautogui.screenshot(region=...) rejects ("must be a tuple of four
        # ints"). Coerce to plain Python ints.
        #
        # The region is anchored to the matched PROGRESS NOTES template. Rather
        # than a single fragile pixel offset to one 40px line (which breaks if
        # the template is recaptured from a different part of the window), scan
        # a GENEROUS band covering the whole area where the patient header can
        # appear - from just below the matched template down ~260px, and wider
        # than the template so the full "Lastname, Firstname (ID) [.. Y].."
        # line is captured. OCR-ing a bigger band and searching it for the ID
        # is more robust than pixel-perfect targeting.
        screen_w, screen_h = pyautogui.size()

        left = int(editor_box.left)
        top = int(editor_box.top) + int(editor_box.height)  # start just below match
        width = max(int(editor_box.width), 520)
        height = 260

        # Clamp to screen bounds so the region is always valid.
        left = max(0, min(left, int(screen_w) - 1))
        top = max(0, min(top, int(screen_h) - 1))
        width = max(1, min(width, int(screen_w) - left))
        height = max(1, min(height, int(screen_h) - top))

        deadline = time.time() + timeout
        last_text = ""
        while time.time() < deadline:
            self._check_abort()
            try:
                last_text = self._read_region_text(left, top, width, height)
            except Exception as exc:
                self.log("error", f"OCR error: {type(exc).__name__}: {exc}")
                last_text = ""

            if self._id_in_text(expected_id, last_text):
                return

            time.sleep(0.5)

        raise RuntimeError(
            f"Could not confirm patient ID {expected_id} in the note header "
            f"within {timeout}s (last OCR read: '{last_text.strip()[:80]}') - "
            "stopping rather than risk pasting into the wrong patient's chart."
        )

    @staticmethod
    def _id_in_text(expected_id, text):
        """
        Match the patient ID against OCR output tolerantly. Tesseract often
        injects spaces inside long digit runs and mangles a char or two, so a
        raw `expected_id in text` substring test throws false negatives. We
        strip every non-digit from BOTH sides and check the clean ID appears
        as a digit-substring. This keeps the safety property (must actually
        see the right ID) while not aborting on cosmetic OCR noise.
        """
        if not text:
            return False
        digits = "".join(c for c in text if c.isdigit())
        return expected_id in digits

    def _clear_template_body(self, editor_box):
        """
        Keep the single header line, delete the template body below it.

        Sends the delete choreography with the dropped-first-keystroke guard
        (Ctrl+Home twice + inter-key delays), then OCR-verifies the body
        actually cleared before returning. If the template sentinel is still
        present, retries the delete once; if it STILL isn't gone, raises so we
        never paste into a half-cleared note (which is what produced the
        "note appended under the intact template" result).
        """

        def _do_delete_sequence():
            # Ctrl+Home TWICE: the first can be dropped by RDP right here, and
            # repeating it is harmless (top stays top). This is the actual fix
            # for the append-below-template bug.
            pyautogui.hotkey("ctrl", "home")
            time.sleep(0.3)
            pyautogui.hotkey("ctrl", "home")
            time.sleep(0.3)
            pyautogui.press("down")          # move off the header onto line 2
            time.sleep(0.2)
            pyautogui.hotkey("ctrl", "shift", "end")  # select body to the end
            time.sleep(0.2)
            pyautogui.press("delete")
            time.sleep(0.4)

        time.sleep(0.8)  # let RDP input settle after the OCR read loop
        _do_delete_sequence()

        if self._body_appears_cleared(editor_box):
            self.log("success", "Template body cleared (verified).")
        else:
            self.log("warn",
                     "Template body still present after delete - retrying the "
                     "clear once (likely another dropped keystroke).")
            _do_delete_sequence()
            if not self._body_appears_cleared(editor_box):
                self._save_timeout_screenshot("body_not_cleared")
                raise RuntimeError(
                    "Could not clear the template body (the "
                    f"'{self.clear_sentinel}' template line is still on screen "
                    "after two attempts) - stopping rather than pasting the "
                    "note under an un-deleted template. A screenshot was saved."
                )

    def _body_appears_cleared(self, editor_box):
        """
        Return True if the template body looks deleted - i.e. the sentinel
        word from the BOTTOM of the template (default 'Educational', the last
        template line) is no longer readable in the editor area.

        Checked BEFORE the note is pasted, so the ONLY possible source of the
        sentinel is leftover template text. Reads a tall region so it reaches
        the bottom of the template.
        """
        try:
            screen_w, screen_h = pyautogui.size()
            left = int(editor_box.left)
            top = int(editor_box.top) + int(editor_box.height)
            width = max(int(editor_box.width), 520)
            height = 460  # tall enough to reach the bottom template line

            left = max(0, min(left, int(screen_w) - 1))
            top = max(0, min(top, int(screen_h) - 1))
            width = max(1, min(width, int(screen_w) - left))
            height = max(1, min(height, int(screen_h) - top))

            text = self._read_region_text(left, top, width, height)
            return self.clear_sentinel.lower() not in text.lower()
        except Exception as exc:
            # If we can't verify, treat as NOT cleared - safer to retry/abort
            # than to assume success and paste into a mess.
            self.log("error", f"Body-clear verify OCR error: {type(exc).__name__}: {exc}")
            return False

    # ---------------- per-patient pipeline ----------------

    def process_patient(self, record):
        self._check_abort()
        self.log("info", f"--- Patient {record.patient_id} ---")

        self._focus_nephro_window()
        time.sleep(1.2)  # increased settle - manual switching being more
                          # reliable than automated switching suggests RDP
                          # needs more time to actually start forwarding input
                          # than a bare GetForegroundWindow check confirms.

        # 1. Open patient search (F3) - with one retry, since RDP sometimes
        #    drops the very first keystroke right after a focus switch.
        pyautogui.press("f3")
        appeared = False
        try:
            self._wait_for_template(
                TEMPLATE_SEARCH_PATIENTS, True, min(15, self.search_open_timeout),
                label="search dialog open"
            )
            appeared = True
        except RuntimeError:
            pass

        if not appeared:
            self.log("warn", "Search dialog didn't appear yet - retrying F3 once.")
            pyautogui.press("f3")
            self._wait_for_template(
                TEMPLATE_SEARCH_PATIENTS, True, self.search_open_timeout,
                label="search dialog open (retry)"
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
        time.sleep(self.post_load_settle)  # dialog closing != patient data fully
                                            # loaded over RDP - extra buffer here
                                            # specifically to avoid landing on the
                                            # PREVIOUS patient's still-displayed screen

        # 4. Ensure the right tab is active (usually already default)
        self._ensure_progress_tab()
        self._check_abort()

        # 5. New Progress Note entry - CLICK the new-note icon anchored to the
        # Progress Notes panel, NOT Ctrl+Insert. The keystroke depends on which
        # pane has focus (it also fires in Doctors Orders) so it "worked
        # sometimes"; the anchored click is focus-independent and Orders-safe.
        # This also acts as an obstruction guard: if the anchored icon isn't
        # visible (panel covered / Nephro not on top), it stops instead of
        # clicking blind. No retry needed - a located click is idempotent in a
        # way a blind keystroke was not, so the old duplicate-note risk is gone.
        self._click_new_note_icon(self.note_editor_timeout)

        # Confirm the floating PROGRESS NOTES editor modal actually opened.
        editor_box = self._wait_for_template(
            TEMPLATE_PROGRESS_NOTES, True, self.note_editor_timeout,
            label="note editor opened"
        )
        self._check_abort()
        time.sleep(0.6)  # settle before reading/pasting

        # 5b. OCR-verify the header actually shows THIS patient's ID before
        # touching anything - replaces guessing a fixed delay with an actual
        # read of what's on screen. Retries automatically; no manual check
        # needed per patient.
        self._wait_for_correct_patient_header(record.patient_id, editor_box)
        self._check_abort()

        # 6. Keep the header line (exactly one line), clear the template body
        # below it, THEN paste - so the note replaces the template instead of
        # being appended under it.
        #
        # Root cause of the earlier "note appended below intact template" bug:
        # over RDP the FIRST keystroke after the OCR read loop gets dropped
        # (the same dropped-first-keystroke issue F3 already guards against).
        # The dropped key was Ctrl+Home, so the caret never went to the top;
        # the selection/delete then did nothing while the later paste still
        # landed - appending. Guards below:
        #   - settle after the OCR loop,
        #   - Ctrl+Home sent TWICE (harmless to repeat; defeats the drop),
        #   - small delays between keys so RDP doesn't coalesce/drop them,
        #   - after deleting, OCR-verify the template body actually cleared
        #     (sentinel word gone) before pasting; retry the delete once, then
        #     abort rather than paste into a half-cleared note.
        self._clear_template_body(editor_box)
        self._check_abort()

        # 7. Paste the clinical note. After the clear, the caret sits at the
        # start of the (now empty) line directly under the header. Add ONE
        # blank line so the note isn't jammed against the header, then paste.
        pyautogui.press("enter")
        time.sleep(0.2)
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
