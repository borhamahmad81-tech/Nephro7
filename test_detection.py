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
import pyautogui

TEMPLATES = [
    "search_patients_dialog.png",
    "progress_notes_editor.png",
    "verification_dialog.png",
    "new_note_icon.png",
]


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
    print("(make sure the relevant Nephro window is visible on screen right now)")
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
    print("Done.")


if __name__ == "__main__":
    main()
