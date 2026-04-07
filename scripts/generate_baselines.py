#!/usr/bin/env python3
"""
generate_baselines.py — Export each .docx fixture to PDF via Microsoft Word.

Requirements:
    - Windows with Microsoft Word installed.
    - Python 3 with the pywin32 package:  pip install pywin32

Usage:
    python scripts/generate_baselines.py

The script walks tests/fixtures/docx/ for every .docx file, opens it in
Word (silently, without UI), and saves a PDF with the same stem into
tests/fixtures/expected/.

Notes:
    - Word must not already be running in the background.
    - The COM automation uses win32com.client which is Windows-only.
    - On CI you would typically commit the generated PDFs so that the
      pipeline can diff against them without needing Word installed.
"""

import os
import sys
from pathlib import Path

try:
    import win32com.client  # type: ignore[import-untyped]
except ImportError:
    print(
        "ERROR: pywin32 is not installed.  Run:  pip install pywin32",
        file=sys.stderr,
    )
    sys.exit(1)

# Resolve project paths relative to this script's location.
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
DOCX_DIR = PROJECT_ROOT / "tests" / "fixtures" / "docx"
EXPECTED_DIR = PROJECT_ROOT / "tests" / "fixtures" / "expected"

# Word constant for "Save As PDF".
WD_FORMAT_PDF = 17


def main() -> None:
    if not DOCX_DIR.is_dir():
        print(f"No fixture directory found at {DOCX_DIR}; nothing to do.")
        return

    docx_files = sorted(DOCX_DIR.glob("*.docx"))
    if not docx_files:
        print(f"No .docx files found in {DOCX_DIR}; nothing to do.")
        return

    EXPECTED_DIR.mkdir(parents=True, exist_ok=True)

    # Start Word as an invisible COM server.
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False  # Suppress any dialogs.

    try:
        for docx_path in docx_files:
            pdf_path = EXPECTED_DIR / f"{docx_path.stem}.pdf"
            print(f"  {docx_path.name}  ->  {pdf_path.name}")

            doc = word.Documents.Open(str(docx_path))
            doc.SaveAs2(str(pdf_path), FileFormat=WD_FORMAT_PDF)
            doc.Close(SaveChanges=False)

        print(f"\nDone — {len(docx_files)} PDF(s) written to {EXPECTED_DIR}")
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
