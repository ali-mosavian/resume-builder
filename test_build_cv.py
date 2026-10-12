# /// script
# requires-python = ">=3.12"
# dependencies = ["python-docx~=1.1", "pyyaml~=6.0", "typer~=0.12"]
# ///
"""Self-check: builds cv_en.yaml with each template and checks what the reader sees. Run: uv run test_build_cv.py"""
import re
import tempfile
from pathlib import Path

from docx import Document
from docx.shared import Cm

from build_cv import build

here = Path(__file__).parent
for template in ("template.docx", "template_pro.docx"):
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "cv.docx"
        build(here / "cv_en.yaml", out, here / template)
        doc = Document(out)
        main = doc.tables[0].rows[0].cells[2]
        texts = [p.text for p in main.paragraphs]

        # Plain text once got an explicit bold=False, which un-bolded every heading.
        assert not any(r.bold is False for p in main.paragraphs for r in p.runs), "explicit bold=False found"
        # **markup** becomes a bold run without the asterisks.
        first = next(p for p in main.paragraphs if "store capacity optimization engine" in p.text)
        assert "**" not in first.text and any(r.bold and r.text == "store capacity optimization engine" for r in first.runs)
        assert "SEP 2025 — PRESENT" in texts
        everywhere = [*doc.paragraphs, *doc.tables[0].rows[0].cells[0].paragraphs]
        assert any(p.text == "Ali Mosavian" for p in everywhere), f"{template}: name placeholder not filled"

# template_pro once had only cell widths, no grid widths: renderers split the page evenly,
# squeezing the main column to a third and doubling the page count.
grid = Document(here / "template_pro.docx").tables[0]._tbl.tblGrid.gridCol_lst
assert grid[2].w > Cm(12), "template_pro main column is not sized in the table grid"
with tempfile.TemporaryDirectory() as tmp:
    from build_html import build as build_html

    for theme in ("pro", "aurora", "aurora-v2"):
        out = Path(tmp) / f"cv_{theme}.html"
        build_html(here / "cv_en.yaml", out, theme)
        html = out.read_text(encoding="utf-8")
        assert "<strong>store capacity optimization engine</strong>" in html and "**" not in html
        assert "reading &amp; learning" in html, "YAML text is not HTML-escaped"
        # Brief roles are not contiguous (Watty sits between them): the caption must appear once.
        assert html.count("data-earlier=") == 1, "earlier-roles caption repeated or missing"
    # One page, both languages: the second starts hidden, the theme script is inlined.
    out = Path(tmp) / "cv_both.html"
    build_html(here / "cv_en.yaml", out, "aurora-v3", also=[here / "cv_sv.yaml"])
    html = out.read_text(encoding="utf-8")
    assert '<div class="page" lang="en"' in html and '<div class="page" lang="sv"' in html
    assert re.search(r'<div class="page" lang="sv" data-ui="[^"]*" hidden>', html), "second language not hidden"
    assert not re.search(r'<div class="page" lang="en" data-ui="[^"]*" hidden>', html), "first language hidden"
    assert html.count("<script>") == 1 and 'data-aka="GCP"' in html
print("ok")
