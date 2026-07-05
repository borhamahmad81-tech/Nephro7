# Medvision Sync Engine

Parses patient progress notes from a Word (.docx) file and enters them into
the Medvision Nephro EMR via keyboard-shortcut automation (F3 search,
Ctrl+Insert new note, Ctrl+Alt+S sign/save), with adaptive window-state
polling instead of fixed delays to handle variable/cloud latency.

## Setup
```
pip install -r requirements.txt
python gui.py
```

## Files
- `docx_parser.py` - extracts {patient_id: note_text} from the source docx
- `automation_engine.py` - drives the EMR via pyautogui/pygetwindow
- `gui.py` - Tkinter/ttk control panel (Slate Dark theme)

## Notes
- Signature password is kept in memory only for the session, never written to disk.
- Guided mode pauses before each sign so you can verify before it commits.
