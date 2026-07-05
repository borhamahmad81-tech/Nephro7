"""
docx_parser.py
Extracts patient records from a clinical progress-notes Word document.

Real file format observed (validated against sample):
    Lastname, Firstname (1089460552) [67 Y] female
    <note text...>
    Lastname, Firstname (1036240040) [55 Y] female
    <note text...>

Header rule: a paragraph is a patient header if it contains a 6-12 digit
number enclosed in parentheses. Everything between one header and the next
(or end of document) belongs to that patient.
"""

import re
from docx import Document

# Matches "(1089460552)" style IDs - 6 to 12 digits inside parentheses.
ID_PATTERN = re.compile(r"\((\d{6,12})\)")


class PatientRecord:
    __slots__ = ("patient_id", "header_text", "note_text")

    def __init__(self, patient_id, header_text, note_lines):
        self.patient_id = patient_id
        self.header_text = header_text.strip()
        # Join note lines with single newlines, strip empty edges,
        # collapse >2 consecutive blank lines to 1.
        text = "\n".join(line for line in note_lines)
        text = re.sub(r"\n{3,}", "\n\n", text).strip()
        self.note_text = text

    def __repr__(self):
        return f"<PatientRecord id={self.patient_id} note_len={len(self.note_text)}>"


def parse_docx(path):
    """
    Parse the given .docx file and return a list of PatientRecord objects,
    in document order.

    Raises:
        ValueError: if no patient headers are found at all (caller should
                    alert the user - "0 patients found").
    """
    doc = Document(path)

    # Pull paragraph texts, skipping fully-empty ones (per spec: ignore
    # empty lines). We keep original paragraph order.
    paragraphs = [p.text for p in doc.paragraphs]

    records = []
    current_id = None
    current_header = None
    current_lines = []

    def flush():
        if current_id is not None:
            records.append(PatientRecord(current_id, current_header, current_lines))

    for para_text in paragraphs:
        stripped = para_text.strip()
        if not stripped:
            continue  # ignore empty lines

        match = ID_PATTERN.search(stripped)
        if match:
            # New patient header found -> close out the previous one.
            flush()
            current_id = match.group(1)
            current_header = stripped
            current_lines = []
        else:
            if current_id is not None:
                current_lines.append(stripped)
            # else: text before the first header in the doc - ignored,
            # since it doesn't belong to any patient.

    flush()  # close out the last patient in the file

    if not records:
        raise ValueError("No patient headers found in document (0 patients).")

    return records


if __name__ == "__main__":
    # Quick manual test hook: python docx_parser.py <file.docx>
    import sys

    target = sys.argv[1] if len(sys.argv) > 1 else None
    if not target:
        print("Usage: python docx_parser.py <path-to-docx>")
        sys.exit(1)

    recs = parse_docx(target)
    print(f"Parsed {len(recs)} patient records.\n")
    for r in recs:
        preview = r.note_text[:80].replace("\n", " ")
        print(f"ID {r.patient_id:>12} | {r.header_text[:45]:<45} | note: {preview}...")
