"""
diagnose.py
Standalone diagnostic - does NOT touch any patient data.

What it does:
  1. Finds and focuses the Nephro/RDP window (same method as the main tool).
  2. Takes a screenshot BEFORE pressing F3, saves it as diag_before_f3.png
  3. Presses F3.
  4. Waits 3 seconds.
  5. Takes a screenshot AFTER, saves it as diag_after_f3.png
  6. Also saves a screenshot of just the Nephro window's title bar area,
     as diag_titlebar_crop.png - useful for building a fresh template.

Run this, then send back the 3 generated PNG files (in the same folder as
this script) so the actual live screen/resolution can be seen directly,
instead of guessing why the template match fails.

Usage: python diagnose.py
"""

import time
import sys

import pyautogui


def find_nephro_hwnd():
    import win32gui

    hwnd = None

    def _enum(h, _):
        nonlocal hwnd
        if hwnd is None and "NEPHRO" in win32gui.GetWindowText(h).upper():
            hwnd = h

    win32gui.EnumWindows(_enum, None)
    return hwnd


def focus_and_settle(hwnd):
    import win32gui
    import win32con
    import ctypes

    if win32gui.GetForegroundWindow() != hwnd:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass
        ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)
        time.sleep(0.3)

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    cx = (left + right) // 2
    cy = top + 12
    pyautogui.click(cx, cy)
    time.sleep(0.2)
    pyautogui.press("escape")
    time.sleep(0.2)
    return (left, top, right, bottom)


def main():
    print("Looking for a window with 'NEPHRO' in its title...")
    hwnd = find_nephro_hwnd()
    if not hwnd:
        print("ERROR: No window with 'NEPHRO' in its title was found.")
        print("Make sure the Nephro/Remote Desktop session is open, then try again.")
        sys.exit(1)

    import win32gui
    title = win32gui.GetWindowText(hwnd)
    print(f"Found window: '{title}'")

    rect = focus_and_settle(hwnd)
    print(f"Window rect: {rect}")
    print("Focused and clicked title bar. Confirming foreground...")

    if win32gui.GetForegroundWindow() != hwnd:
        print("WARNING: Could not confirm foreground focus after settling.")
    else:
        print("Confirmed: Nephro window has OS foreground focus.")

    print("Taking BEFORE screenshot...")
    before = pyautogui.screenshot()
    before.save("diag_before_f3.png")

    left, top, right, bottom = rect
    titlebar_crop = before.crop((left, top, right, top + 30))
    titlebar_crop.save("diag_titlebar_crop.png")

    print("Pressing F3 now...")
    pyautogui.press("f3")

    print("Waiting 3 seconds...")
    time.sleep(3)

    print("Taking AFTER screenshot...")
    after = pyautogui.screenshot()
    after.save("diag_after_f3.png")

    print()
    print("Done. Please send back these 3 files from this folder:")
    print("  - diag_before_f3.png")
    print("  - diag_after_f3.png")
    print("  - diag_titlebar_crop.png")


if __name__ == "__main__":
    main()
