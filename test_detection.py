"""
test_detection.py
Standalone self-test - does NOT touch any patient data or press any keys.

Checks two separate things that both look identical from the outside
("template not found") but have very different fixes:

  1. Is OpenCV/image-matching actually working at all in this environment?
  2. If it IS working, do the 3 template images match what's currently
     on screen (with whatever window is open right now)?

Usage: run this with whatever Nephro window/dialog you want to test
already open on screen, then read the printed results.
"""

import sys
import time
import pyautogui

TEMPLATES = [
    "search_patients_dialog.png",
    "progress_notes_editor.png",
    "verification_dialog.png",
    "new_note_icon.png",
]


def _minimize_self_console():
    """
    Minimize THIS console window automatically. Windows activates whatever
    window was on top before this one when it's minimized - normally the
    Nephro window the user had open right before double-clicking this EXE.
    This replaces relying on the user to Alt+Tab within a countdown, which
    is exactly the kind of manual-timing step this project keeps hitting
    problems with.
    """
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            SW_MINIMIZE = 6
            ctypes.windll.user32.ShowWindow(hwnd, SW_MINIMIZE)
    except Exception:
        pass  # non-fatal - falls back to whatever is currently on screen


def _restore_self_console():
    """Bring this console window back so the results can be read."""
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            SW_RESTORE = 9
            ctypes.windll.user32.ShowWindow(hwnd, SW_RESTORE)
            ctypes.windll.user32.SetForegroundWindow(hwnd)
    except Exception:
        pass


def test_ocr():
    print("=== Step 0: checking if OCR (Tesseract) is working ===")
    try:
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
                print(f"Using bundled Tesseract at: {exe_path}")
            else:
                print(f"WARNING: expected bundled Tesseract not found at {exe_path}")

        shot = pyautogui.screenshot()
        text = pytesseract.image_to_string(shot)
        print(f"OCR OK - read {len(text)} characters from current screen.")
        print(f"Sample of what it read: {text.strip()[:150]!r}")
    except Exception as exc:
        print(f"OCR FAILED: {type(exc).__name__}: {exc}")
    print()


def main():
    print("=== Nephro detection self-test ===")
    print()
    print("This window will minimize itself automatically in a couple of")
    print("seconds so the window it's covering (expected: Nephro) is what")
    print("gets captured - not this console window. It restores itself when")
    print("done. You don't need to Alt+Tab or click anything.")
    print()
    time.sleep(2)

    _minimize_self_console()
    time.sleep(1.5)  # let Windows finish activating whatever is now on top

    test_ocr()

    print("=== Step 1: checking if image matching works at all ===")
    try:
        import cv2
        print(f"OpenCV import OK - version {cv2.__version__}")
    except Exception as exc:
        print(f"FAILED to import cv2 directly: {type(exc).__name__}: {exc}")
        print("This means OpenCV isn't properly bundled/installed - that alone")
        print("would explain every template match failing regardless of content.")

    print()
    print("=== Step 2: testing each template against the CURRENT screen ===")
    print()

    for template in TEMPLATES:
        try:
            box = pyautogui.locateOnScreen(template, confidence=0.8)
            if box:
                print(f"[FOUND]   {template} -> matched at {box}")
            else:
                print(f"[NOT FOUND] {template} -> no match on current screen")
        except pyautogui.ImageNotFoundException:
            print(f"[NOT FOUND] {template} -> no match on current screen")
        except Exception as exc:
            print(f"[ERROR]   {template} -> {type(exc).__name__}: {exc}")

    print()
    print("Also saving a full screenshot as detection_screen.png for reference.")
    pyautogui.screenshot().save("detection_screen.png")

    _restore_self_console()

    print()
    print("Done - this window restored itself. Results are above.")
    print()
    input("Press Enter to close this window...")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        _restore_self_console()
        print()
        print(f"CRASHED: {type(exc).__name__}: {exc}")
        input("Press Enter to close this window...")
