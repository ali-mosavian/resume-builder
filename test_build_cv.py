# /// script
# requires-python = ">=3.12"
# dependencies = ["python-docx~=1.1", "pyyaml~=6.0", "typer~=0.12"]
# ///
"""Self-check: builds cv_en.yaml and checks what the reader sees. Run: uv run test_build_cv.py"""
import tempfile
from pathlib import Path

from docx import Document

from build_cv import build

here = Path(__file__).parent
with tempfile.TemporaryDirectory() as tmp:
    out = Path(tmp) / "cv.docx"
    build(here / "cv_en.yaml", out)
    main = Document(out).tables[0].rows[0].cells[2]
    texts = [p.text for p in main.paragraphs]

    # Plain text once got an explicit bold=False, which un-bolded every heading.
    assert not any(r.bold is False for p in main.paragraphs for r in p.runs), "explicit bold=False found"
    # **markup** becomes a bold run without the asterisks.
    first = next(p for p in main.paragraphs if "store capacity optimization engine" in p.text)
    assert "**" not in first.text and any(r.bold and r.text == "store capacity optimization engine" for r in first.runs)
    assert "SEP 2025 — PRESENT" in texts
print("ok")
