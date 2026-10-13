"""templates/pro.docx: name banner on top, light sidebar left, main column right."""

from functools import partial
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm

from resume.docx_templates import Mark, border, bullet_lists, cell, style, table

FONT = "Calibri"
INK, ACCENT, MUTED, RULE, FILL = "222222", "1F3A5F", "6B7280", "C9D1DB", "F3F5F8"


def make(out: Path) -> None:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.6)
    sec.top_margin = sec.bottom_margin = Cm(1.4)

    font = partial(style, doc, font=FONT, color=INK)
    font("Normal", size=9.5, after=3)
    font("Name", "Normal", size=24, color=ACCENT, bold=True, after=0, line=1.0)
    border(font("Job Title", "Normal", size=10, color=MUTED, spacing=1.5, after=10), "bottom", ACCENT, 12, 2)
    h1 = font("Heading 1", "Normal", size=10.5, color=ACCENT, bold=True, caps=True, spacing=1, before=4, after=5)
    border(h1, "bottom", RULE, 6, 2)
    font("Heading 2", "Normal", size=10.5, bold=True, before=6, after=0)
    font("Heading 3", "Normal", size=9, color=ACCENT, bold=True, caps=True, spacing=1, after=3)
    font("Heading 4", "Normal", size=7.5, color=MUTED, caps=True, spacing=0.8, before=3, after=0)
    font("Date", "Normal", size=8, color=MUTED, spacing=0.5, after=3)
    font("Sidebar Text", "Normal", size=9, after=1)
    font("List Paragraph", "Normal", size=9.5, after=2)
    font("Sidebar Sections Spacing", "Normal", size=4, after=6)
    font("Main Sections Spacing", "Normal", size=4, after=4)
    bullet_lists(doc, [Mark("•", ACCENT, 284, 227)])

    doc.add_paragraph(style="Name").add_run("{Name}")
    doc.add_paragraph(style="Job Title").add_run("{Job Title}")

    widths = (5.0, 0.5, 12.3)  # sums to the 17.8 cm text width
    side, gap, body = table(doc, widths, align=WD_TABLE_ALIGNMENT.CENTER).rows[0].cells
    cell(side, widths[0], fill=FILL, pad=(0.35, 0.35, 0.35, 0.35))
    cell(gap, widths[1])
    cell(body, widths[2])
    side.paragraphs[0]._p.style = "SidebarSectionsSpacing"
    doc.save(str(out))
