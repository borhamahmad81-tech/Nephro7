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
        # Sentinel/marker settings used when verifying the body was cleared.
        # 'Complaint' is the FIRST body line, so if it is absent the whole
        # template body is gone. It must not appear in the one-line header.
        self.clear_sentinel = "Complaint"

        # Vertical offset (pixels) from the TOP of the matched PROGRESS NOTES
        # tab template down to the editor BODY TEXT (around the "Complaint"
        # line), where we click to lock input focus before clearing. An earlier
        # value landed on the toolbar, so focus never reached the document and
        # no keystroke afterwards did anything. The editor opens at a
        # consistent vertical spot, so a fixed offset works. If the focus-click
        # misses the text, adjust: larger = lower.
        self.editor_click_dy = 190

        # Vertical offset (pixels) from the TOP of the matched PROGRESS NOTES
        # tab template down to the start of the editor's TEXT AREA - i.e. past
        # the toolbar, the Date/Time/Doctor row, the format row and the ruler.
        #
        # This exists because OCR checks were previously reading from just
        # below the matched tab, which is the TOOLBAR: they always picked up
        # ~26 characters of "Date 07/23/2026 Time ... Doctor" text, so the
        # "is the editor empty?" check could never return True even when the
        # note was genuinely blank - and the run aborted before pasting.
        #
        # Calibrated against the live layout: ruler sits ~117px below the tab
        # top, first line of text ~137px. 130 clears the ruler and starts just
        # above the first text line. If OCR ever reads toolbar text again,
        # increase this; if it misses the first line of the note, decrease it.
        self.text_area_dy = 130

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
        # Park the cursor centre-screen BEFORE any pyautogui action in this
        # whole per-patient cycle. The fail-safe fires on the NEXT pyautogui
        # call if the cursor is sitting in a corner, so if the user left the
        # mouse in a corner (or a previous step nudged it there) the very first
        # click/press aborts. Doing this here, at the top of the cycle, covers
        # everything downstream. SetCursorPos does not go through the fail-safe.
        self._park_cursor_safe()

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

        for attempt in range(2):
            try:
                # Re-park the cursor centre-screen IMMEDIATELY before the click
                # every attempt. The fail-safe checks the cursor position at the
                # moment of the pyautogui call, so parking must be the last
                # thing before it.
                self._park_cursor_safe()
                time.sleep(0.05)
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                cx = (left + right) // 2
                # Click a point safely INSIDE the title bar strip, clamped away
                # from the very top edge so the move can never read as a corner.
                cy = max(top + 12, 8)
                pyautogui.click(cx, cy)
                time.sleep(0.2)
                pyautogui.press("escape")  # clear any accidentally-opened menu
                time.sleep(0.2)
                return
            except pyautogui.FailSafeException:
                # The cursor was in a corner at the instant of the click. Park
                # and retry once rather than abandoning focus settling (which
                # would leave keystrokes going into the void).
                self.log("warn",
                         "Fail-safe tripped during focus settle - re-parking "
                         "cursor and retrying once.")
                self._park_cursor_safe()
                time.sleep(0.2)
            except Exception as exc:
                self.log("warn", f"Could not settle RDP input focus: {exc}")
                return

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

    def _wait_for_patient_in_notes_panel(self, expected_id, timeout=None):
        """
        Wait until the DOCKED Progress Notes panel shows the expected patient's
        ID, i.e. the new patient's notes have actually finished loading over
        RDP - before we open a new note. This is the real "safe to proceed"
        signal, unlike the "Administration no." field which updates too early.

        Anchored to the new-note-icon template (which lives in the docked
        panel's toolbar): once found, we OCR a region spanning the panel's
        content area just below it and look for the expected ID. If the panel
        still shows the PREVIOUS patient (different ID), we keep waiting; if the
        new ID never appears within the timeout, we STOP rather than risk
        opening a note against a half-loaded/previous patient.
        """
        if timeout is None:
            timeout = self.patient_load_timeout

        deadline = time.time() + timeout
        last_text = ""
        while time.time() < deadline:
            self._check_abort()

            # Locate the docked panel via the new-note icon anchor (it sits in
            # that panel's toolbar). This makes the OCR region track the window.
            anchor = self._find_on_screen(TEMPLATE_NEW_NOTE_ICON)
            if anchor is None:
                # Panel/icon not visible yet (still loading or covered) - wait.
                time.sleep(0.5)
                continue

            try:
                screen_w, screen_h = pyautogui.size()
                # OCR a band starting just below the icon/toolbar, covering the
                # panel's header/content lines where "Lastname, Firstname (ID)"
                # appears.
                left = int(anchor.left) - 10
                top = int(anchor.top) + int(anchor.height) + 20
                width = 520
                height = 120

                left = max(0, min(left, int(screen_w) - 1))
                top = max(0, min(top, int(screen_h) - 1))
                width = max(1, min(width, int(screen_w) - left))
                height = max(1, min(height, int(screen_h) - top))

                last_text = self._read_region_text(left, top, width, height)
                digits = "".join(c for c in last_text if c.isdigit())
                if expected_id in digits:
                    self.log("success",
                             f"Progress Notes panel shows patient {expected_id} "
                             "- fully loaded, safe to open a note.")
                    return
            except Exception as exc:
                self.log("error", f"Notes-panel OCR error: {type(exc).__name__}: {exc}")

            time.sleep(0.5)

        self._save_timeout_screenshot("patient_not_loaded_in_panel")
        raise RuntimeError(
            f"The Progress Notes panel never showed patient {expected_id} "
            f"within {timeout}s (last OCR: '{last_text.strip()[:60]}') - "
            "stopping rather than opening a note against the wrong or "
            "half-loaded patient. A screenshot was saved."
        )

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

    def _select_all_and_replace(self, editor_box, record):
        """
        Replace the entire note content: Ctrl+A -> Delete -> paste the docx
        header followed by the clinical note.

        WHY THIS METHOD. Through this pipeline, two-modifier combos
        (Ctrl+Shift+End) do not work - proven by a clean experiment: the two
        Enters inserted just before it DID land (a visible gap appeared under
        the header), so focus was live and every plain key and single-modifier
        combo worked; only the selection failed. That rules out every approach
        that tries to select just the body below the header.

        Ctrl+A is a SINGLE modifier combo - the same class as Ctrl+Home, which
        demonstrably works here - so it has a far better chance. It selects
        everything INCLUDING the header, which is acceptable because we paste
        our own header back from the docx (record.header_text).

        TRADE-OFF: the pasted header is plain text and loses the EMR's bold
        styling. This is a deliberate trade for a clear that actually works.

        SAFETY:
          - The correct-patient check (step 5b) already confirmed the EMR's own
            header showed THIS patient BEFORE anything is deleted.
          - After Delete we OCR-verify the editor is genuinely EMPTY. If any
            text remains, the select/delete failed - we retry once, then abort
            rather than paste on top of a surviving template.
          - After pasting we OCR-verify the patient ID is present again, so a
            failed paste can't silently leave an empty note.
        """
        MOD_DELAY = 0.18

        def _ctrl_a_delete():
            # Ctrl+A with explicit hold + delays (single modifier).
            pyautogui.keyDown("ctrl")
            time.sleep(MOD_DELAY)
            pyautogui.keyDown("a")
            time.sleep(MOD_DELAY)
            pyautogui.keyUp("a")
            time.sleep(MOD_DELAY)
            pyautogui.keyUp("ctrl")
            time.sleep(0.3)
            pyautogui.press("delete")
            time.sleep(0.5)

        def _focus_click():
            self._park_cursor_safe()
            bx, by = int(editor_box.left), int(editor_box.top)
            focus_x = bx + 60
            focus_y = by + self.editor_click_dy
            self.log("info",
                     f"Clicking editor body to lock focus at ({focus_x}, {focus_y}).")
            pyautogui.click(focus_x, focus_y)
            time.sleep(0.5)

        # --- clear everything ---
        _focus_click()
        self.log("info", "Selecting all (Ctrl+A) and deleting...")
        _ctrl_a_delete()

        if not self._editor_appears_empty(editor_box):
            self.log("warn",
                     "Editor not empty after Ctrl+A + Delete - retrying once.")
            _focus_click()
            _ctrl_a_delete()
            if not self._editor_appears_empty(editor_box):
                self._save_timeout_screenshot("editor_not_empty")
                raise RuntimeError(
                    "Ctrl+A + Delete did not clear the note after two attempts "
                    "- stopping rather than pasting on top of the existing "
                    "template. A screenshot was saved."
                )

        self.log("success", "Editor cleared (verified empty).")

        # --- paste header + note ---
        header = (record.header_text or "").strip()
        body = record.note_text or ""
        content = (header + "\n\n" + body) if header else body

        pyperclip.copy(content)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(1.0)

        # Scroll back to the TOP before verifying. After pasting a long note
        # the editor view scrolls to the bottom (the caret ends at the end of
        # the pasted text), so the header line is no longer on screen and OCR
        # cannot see the patient ID - which made verification fail even though
        # the paste had worked perfectly. Ctrl+Home is a single-modifier combo,
        # the class that does work through this pipeline.
        pyautogui.keyDown("ctrl")
        time.sleep(MOD_DELAY)
        pyautogui.keyDown("home")
        time.sleep(MOD_DELAY)
        pyautogui.keyUp("home")
        time.sleep(MOD_DELAY)
        pyautogui.keyUp("ctrl")
        time.sleep(0.6)  # let the view finish scrolling before reading

        # Verify the paste actually landed and shows this patient.
        header_present, _ = self._read_clear_state(editor_box, record.patient_id)
        if not header_present:
            time.sleep(0.6)
            header_present, _ = self._read_clear_state(editor_box, record.patient_id)
        if not header_present:
            self._save_timeout_screenshot("paste_not_verified")
            raise RuntimeError(
                f"After pasting, patient ID {record.patient_id} was not visible "
                "in the note - the paste may have failed. Stopping rather than "
                "signing an empty or wrong note. A screenshot was saved."
            )

        self.log("success",
                 f"Header + note pasted and verified for patient "
                 f"{record.patient_id}.")

    def _text_area_region(self, editor_box):
        """
        Return (left, top, width, height) for the editor's TEXT AREA only -
        excluding the toolbar above it.

        This fixes a real bug: the region used to start at
        editor_box.top + editor_box.height, i.e. immediately below the matched
        PROGRESS NOTES tab template - which lands on the TOOLBAR (the
        "Date / Time / Doctor" row). OCR therefore always read about 26
        characters of toolbar text, so the "is the editor empty?" check could
        never return True even when the note was completely blank, and the run
        aborted before pasting.

        We now offset down by self.text_area_dy, which is measured to clear the
        toolbar/format/ruler rows and start at the first line of actual text.
        """
        screen_w, screen_h = pyautogui.size()

        left = int(editor_box.left)
        top = int(editor_box.top) + int(self.text_area_dy)
        width = max(int(editor_box.width), 520)
        height = 420

        left = max(0, min(left, int(screen_w) - 1))
        top = max(0, min(top, int(screen_h) - 1))
        width = max(1, min(width, int(screen_w) - left))
        height = max(1, min(height, int(screen_h) - top))
        return left, top, width, height

    def _editor_appears_empty(self, editor_box):
        """
        True if the editor's text area reads as essentially empty via OCR.

        Used right after Ctrl+A + Delete. Because that deletes the header too,
        "empty" here means NO meaningful text at all - not "header only". A
        couple of stray characters are tolerated as OCR noise.
        """
        try:
            left, top, width, height = self._text_area_region(editor_box)
            text = self._read_region_text(left, top, width, height)
            stripped = "".join(c for c in text if c.isalnum())
            self.log("info",
                     f"Empty check: {len(stripped)} alphanumeric chars remain. "
                     f"OCR sample: {text.strip()[:60]!r}")
            return len(stripped) <= 2
        except Exception as exc:
            # Can't verify -> treat as NOT empty, so we retry/abort rather than
            # assume success and paste into a surviving template.
            self.log("error", f"Empty-check OCR error: {type(exc).__name__}: {exc}")
            return False

    def _read_clear_state(self, editor_box, expected_id):
        """
        OCR the editor region once and report two booleans:
          (header_present, body_present)

        header_present: the expected patient ID digits are still visible (the
                        bold header line is intact).
        body_present:   True if meaningful template/body text remains BELOW
                        the header line.

        FIX: this used to check only "does the word 'Complaint' still appear?"
        But a PARTIAL deletion (backspacing "Complaint" down to just "Co")
        already makes that exact word not match, so the old check declared
        "body gone" while a remnant ("Co") was still on screen and got pasted
        over. Now we isolate the text BELOW the header line (splitting on the
        header, identified by the patient ID digits) and check whether
        anything substantial remains there at all - not just one keyword.
        Bullet characters, whitespace, and stray punctuation are stripped
        before judging "substantial", so a lone leftover bullet doesn't count,
        but a word fragment like "Co" correctly does.

        Used by the clear routine to decide: done, keep going, or abort.
        (body gone), or abort (header gone = overshoot).
        """
        try:
            left, top, width, height = self._text_area_region(editor_box)
            text = self._read_region_text(left, top, width, height)
            digits = "".join(c for c in text if c.isdigit())
            header_present = expected_id in digits

            # Isolate text BELOW the header line: split at the line containing
            # the patient ID, keep only what comes after it.
            lines = text.splitlines()
            below_header_lines = []
            header_line_seen = False
            for line in lines:
                if not header_line_seen:
                    if expected_id in "".join(c for c in line if c.isdigit()):
                        header_line_seen = True
                    continue
                below_header_lines.append(line)
            below_text = "\n".join(below_header_lines) if header_line_seen else text

            # Strip bullets/whitespace/punctuation noise; anything substantial
            # left over means the body isn't cleared yet.
            stripped = "".join(
                c for c in below_text if c.isalnum()
            ).strip()
            body_present = len(stripped) > 2  # tolerate a stray OCR artifact char

            return header_present, body_present
        except Exception as exc:
            # If OCR fails, report header MISSING + body PRESENT: the safest
            # combination, because it makes the loop treat it as a potential
            # overshoot and abort rather than keep backspacing blindly.
            self.log("error", f"Clear-state OCR error: {type(exc).__name__}: {exc}")
            return False, True

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

        # 2. Enter the ID via CLIPBOARD PASTE, not typing. Typed characters go
        # through the RDP session's keyboard layout, which on this system can
        # produce the wrong characters (e.g. Greek letters). Pasting sends the
        # actual text and bypasses the layout entirely. Then Enter to search,
        # Enter to confirm/open the result.
        pyperclip.copy(record.patient_id)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.3)
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

        # 3b. Wait until the DOCKED Progress Notes panel actually shows THIS
        # patient's ID before doing anything. The "Administration no." field
        # top-right updates EARLY (before the notes finish loading), so it is
        # NOT a safe signal - checking it would let us proceed while the panel
        # still shows the previous patient. The docked panel's own content is
        # the real "fully loaded" signal, so we OCR that region and wait for
        # the new ID. Replaces the old fixed post_load_settle guess.
        self._wait_for_patient_in_notes_panel(record.patient_id)
        self._check_abort()

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

        # 6+7. Replace the ENTIRE note content: Ctrl+A -> Delete -> paste the
        # header (from the docx) followed by the clinical note.
        #
        # Why this approach: through this pipeline, two-modifier combos like
        # Ctrl+Shift+End do not work (confirmed - the plain keys around it all
        # landed, only the selection failed), so no method that tries to select
        # just the body below the header is reliable. Ctrl+A is a SINGLE
        # modifier combo, the same class as Ctrl+Home which does work here, so
        # it has a much better chance. It selects everything including the
        # header, which is fine because we paste our own header back.
        #
        # TRADE-OFF: the pasted header is plain text and loses the EMR's bold
        # formatting. Accepted deliberately in exchange for a clear that works.
        #
        # Safety: the correct-patient check above (5b) already confirmed the
        # EMR's own header showed THIS patient BEFORE anything is deleted, and
        # after the delete we verify the editor is genuinely empty before
        # pasting - so we can't paste on top of a failed clear.
        self._select_all_and_replace(editor_box, record)
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

            # CRITICAL: the guided popup is a window of OUR OWN GUI. While it
            # was up, and after the user clicks Continue, keyboard focus is on
            # the GUI - NOT on Nephro. Without re-focusing here, the password
            # below would be typed/pasted into the wrong window entirely.
            self.log("info", "Re-focusing Nephro after the guided pause...")
            self._focus_nephro_window()
            time.sleep(0.8)
            self._check_abort()

            # Re-confirm the verification dialog is still the thing on screen
            # before sending credentials into it.
            self._wait_for_template(
                TEMPLATE_VERIFICATION, True, self.verify_open_timeout,
                label="verification dialog still open after guided pause"
            )
            self._check_abort()

        # 9. Enter the password by TYPING it.
        #
        # TYPING IS PRIMARY, DELIBERATELY. An earlier revision changed this to
        # clipboard paste (to dodge the Greek-letter keyboard-layout problem
        # seen in the note editor). That was a mistake: password fields in this
        # app block Ctrl+V, so the paste silently entered nothing and every
        # sign-off failed - a step that had previously worked when typed.
        # Typing is therefore restored as the primary method.
        #
        # (The Greek-letter issue does not apply here in practice: it appeared
        # when typing into a "cold" editor, and by this point a paste has
        # already happened in the session.)
        #
        # Paste is kept only as a FALLBACK, tried once if typing didn't close
        # the dialog.
        self.log("info", "Entering signature password (typing)...")
        pyautogui.typewrite(self.password, interval=0.05)
        time.sleep(0.4)
        pyautogui.press("enter")
        time.sleep(1.5)

        if self._find_on_screen(TEMPLATE_VERIFICATION) is not None:
            self.log("warn",
                     "Verification dialog still open after typing the password "
                     "- trying clipboard paste as a fallback.")
            pyperclip.copy(self.password)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.4)
            pyautogui.press("enter")
            time.sleep(1.0)
            try:
                pyperclip.copy("")  # don't leave the password in the clipboard
            except Exception:
                pass

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
