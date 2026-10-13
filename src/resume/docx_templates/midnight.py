"""templates/midnight.docx, the Word twin of themes/midnight.css.

Dark page: full-page rectangles in the headers draw the night background, the headline cards and the sidebar panel.
"""

from typing import Any
from pathlib import Path

from docx import Document
from docx.shared import Cm
from docx.document import Document as Doc
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.styles.style import ParagraphStyle
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE

from resume.docx_templates import Mark
from resume.docx_templates import cell
from resume.docx_templates import rect
from resume.docx_templates import style
from resume.docx_templates import table
from resume.docx_templates import border
from resume.docx_templates import finish
from resume.docx_templates import numbered
from resume.docx_templates import char_style
from resume.docx_templates import bullet_lists

SANS, MONO = "Inter", "JetBrains Mono"
BG, PANEL, CHIP, EDGE, TRACK = "07080D", "0F1015", "181A23", "2A2B34", "22232C"
INK, INK2, MUTED, FAINT, INDIGO, CYAN = "EEF0F6", "C3C8D6", "9AA2B8", "6E768C", "8B8CF8", "5FE3F2"
GAP = {SANS: 1.21, MONO: 1.32}  # line height per em: Word's "multiple" spacing scales it, CSS line-height doesn't

# Page geometry in mm, as themes/midnight.css lays out an A4 sheet.
CARD_Y, CARD_H, BODY_Y = 42, 20, 69  # headline cards; where the body and the sidebar panel start on page 1
SIDE_X, SIDE_W = 14, 52
HEAD_SKEW = 1  # the first-page header paragraph pushes the body down about this much


def pt(mm: float) -> float:
    return mm / 25.4 * 72


def styles(doc: Doc) -> None:
    def font(name: str, face: str = SANS, **kw: Any) -> ParagraphStyle:
        kw.setdefault("color", INK2)
        kw.setdefault("line", 1.55)
        kw.setdefault("base", None if name == "Normal" else "Normal")
        return style(doc, name, font=face, line_gap=GAP[face], **kw)

    # One list per style: Word links a list to a single style and drops the marks of any other.
    _, diamond, ring = bullet_lists(  # word.py's bullets take the first
        doc,
        [
            Mark("•", INDIGO, 624, 240, 10, "Segoe UI Symbol"),
            Mark("◆", CYAN, 283, 283, 6, "Segoe UI Symbol"),
            Mark("○", INDIGO, 340, 340, 10, "Segoe UI Symbol"),
        ],
    )

    font("Normal", size=9, after=pt(3))
    font("Name", size=34, bold=True, color=INK, spacing=-1.5, after=0, line=1.0)
    font("Job Title", size=9, color=INDIGO, caps=True, spacing=0.9, before=pt(3), after=0, line=1.2)
    font("CV Photo", size=2, after=0, line=1.0)
    contact = font("CV Contact", size=8, bold=True, color=INK, after=0, line=1.2)
    contact.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    border(contact, "top between bottom", EDGE, 4, 5)
    font("CV Highlights", size=2, after=0)
    value = font("CV Highlight Value", size=17, bold=True, color=INK, before=pt(4), after=0, line=1.0)
    label = font("CV Highlight Label", size=7.5, color=MUTED, before=pt(1.4), after=0, line=1.3)
    value.paragraph_format.left_indent = label.paragraph_format.left_indent = Cm(0.36)

    numbered(
        font("Heading 1", MONO, size=7.6, bold=True, color=MUTED, caps=True, spacing=1.2, after=pt(4), line=1.2),
        diamond,
    )
    font("Heading 2", size=9.7, bold=True, color=INK, after=0, line=1.2)
    numbered(font("CV Role", base="Heading 2", size=9.7, bold=True, color=INK, after=0, line=1.2), ring)
    font("CV Company", size=8.2, color=MUTED, after=0, line=1.42).paragraph_format.left_indent = Cm(0.6)
    font("Date", MONO, size=7.2, color=FAINT, spacing=0.3, before=pt(1), after=pt(2), line=1.4)
    font(
        "CV Role Date", MONO, size=7.2, color=FAINT, spacing=0.3, before=pt(1), after=pt(2), line=1.4
    ).paragraph_format.left_indent = Cm(0.6)
    font("Heading 3", MONO, size=7.4, bold=True, color=MUTED, caps=True, spacing=1.2, after=pt(3), line=1.2)
    font(
        "Heading 4",
        MONO,
        size=6.6,
        bold=True,
        color=FAINT,
        caps=True,
        spacing=0.8,
        before=pt(2.4),
        after=pt(0.3),
        line=1.4,
    )
    font("Sidebar Text", size=8.5, after=1)
    kv = font("CV Key Value", size=8.5, after=3)
    kv.paragraph_format.tab_stops.add_tab_stop(Cm(4.2), alignment=WD_TAB_ALIGNMENT.RIGHT)  # at the sidebar text edge
    font("CV Skill", size=8.5, after=pt(1))
    font("CV Skill Bar", size=1, after=3.4, line=1.0)
    font("List Paragraph", size=9, after=pt(1)).paragraph_format.left_indent = Cm(0.6)
    font("CV Tools", MONO, size=7, before=pt(3), after=6, line=2.46).paragraph_format.left_indent = Cm(0.6)
    quote = font("CV Quote", size=9, after=3)
    quote.font.italic = True
    quote.paragraph_format.left_indent = Cm(0.48)
    border(quote, "left", INDIGO, 18, 11)
    font("Sidebar Sections Spacing", size=4, after=4.8)
    font("Main Sections Spacing", size=4, after=10)
    char_style(doc, "CV Accent", font=SANS, color=CYAN)
    char_style(doc, "CV Muted", font=SANS, color=MUTED, size=7.5)
    char_style(doc, "CV Tool", font=MONO, color=INK2, fill=CHIP, size=7)
    char_style(doc, "CV Tool Mark", font=MONO, fill=CYAN, size=7)  # a chip's cyan left edge
    char_style(doc, "CV Skill Fill", font=SANS, fill=INDIGO)
    char_style(doc, "CV Skill Track", font=SANS, fill=TRACK)
    for s in ("CV Contact", "CV Role", "CV Company", "Date", "CV Role Date", "Heading 1", "Heading 2", "Heading 3"):
        doc.styles[s].paragraph_format.keep_with_next = True


def header(doc: Doc) -> None:
    """Two fixed rows: name and links, then the headline numbers over their card shapes."""
    head = table(doc, [15.0, 6.0], rows=2)
    for row, mm in zip(head.rows, (CARD_Y - HEAD_SKEW, CARD_H), strict=True):
        row.height, row.height_rule = Cm(mm / 10), WD_ROW_HEIGHT_RULE.EXACTLY
    name_c, contact_c = head.rows[0].cells
    cell(name_c, 15.0, pad=(1.5, 0, 0, 1.4))
    cell(contact_c, 6.0, pad=(1.5, 1.4, 0, 0))
    name_c.paragraphs[0]._p.style = "Name"
    name_c.paragraphs[0].add_run("{Name}")
    name_c.add_paragraph(style="JobTitle").add_run("{Job Title}")
    contact_c.paragraphs[0]._p.style = "CVContact"
    contact_c.paragraphs[0].add_run("{Contact}")
    stats = head.rows[1].cells[0].merge(head.rows[1].cells[1])
    cell(stats, 21, pad=(0, 1.25, 0, 1.25))
    stats.paragraphs[0]._p.style = "CVHighlights"
    stats.paragraphs[0].add_run("{Highlights}")


def layout(doc: Doc) -> None:
    """Body: main column and sidebar column (its panel is a shape, see decorate)."""
    doc.add_paragraph(style="CV Photo")  # keeps Word from merging the two tables
    side, main = table(doc, [7.8, 13.2]).rows[0].cells
    cell(side, 7.8, pad=(0.37, 1.7, 0, 1.9))
    cell(main, 13.2, pad=(0.62, 1.4, 0, 0))
    side.paragraphs[0]._p.style = "CVPhoto"  # word.py adds a section spacer of its own
    doc.add_paragraph(style="CV Photo")  # Word needs a paragraph between the last table and the section end


def panel(p: Paragraph, box: tuple[float, float, float, float], shape_id: int) -> None:
    """A filled box with a 0.25 mm edge."""
    x, y, w, h = box
    rect(p, (x, y, w, h, EDGE), shape_id)
    rect(p, (x + 0.25, y + 0.25, w - 0.5, h - 0.5, PANEL), shape_id + 1)


def decorate(doc: Doc) -> None:
    """Shapes behind the text, repeated by the page headers. Page 1: night, the four cards, the sidebar panel.
    Later pages: night only, and the body starts 12 mm down, as the web's later sheets do."""
    sec = doc.sections[0]
    first = sec.first_page_header.paragraphs[0]
    first._p.style = "CVPhoto"
    rect(first, (0, 0, 210, 297, BG), 1001)
    for i in range(4):
        panel(first, (12.5 + 46.25 * i + 1.5, CARD_Y, 43.25, CARD_H), 1010 + 2 * i)
    panel(first, (SIDE_X, BODY_Y, SIDE_W, 287 - BODY_Y), 1030)
    later = sec.header.paragraphs[0]
    later._p.style = "CVPhoto"
    later.paragraph_format.space_before = Cm(0.7)
    rect(later, (0, 0, 210, 297, BG), 1101)


def make(out: Path) -> None:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(0)
    sec.top_margin, sec.bottom_margin = Cm(0), Cm(1.0)
    sec.header_distance, sec.footer_distance = Cm(0), Cm(0)
    sec.different_first_page_header_footer = True
    styles(doc)
    header(doc)
    layout(doc)
    decorate(doc)
    finish(doc, out, SANS, MONO)
