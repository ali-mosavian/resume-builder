"""templates/aurora.docx, the Word twin of themes/aurora.css.

Opts into word.py's optional parts: Contact and Highlights placeholders (no photo), skill bars, a role
timeline, tool chips and quote bars.
"""

from pathlib import Path
from functools import partial

from docx import Document
from docx.shared import Cm
from docx.document import Document as Doc
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

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

SANS, SERIF = "Inter", "Fraunces"
SANS_GAP, SERIF_GAP = 1.21, 1.233  # line height per em: Word's "multiple" spacing scales it, CSS line-height doesn't
INK, MUTED, VIOLET, TEAL, NIGHT = "151A2D", "667085", "7C5CFF", "19C3B1", "0B1024"
LAVENDER, LINE, CHIP, CHIP_INK, RULE = "A5B4FC", "E4E7EF", "EEF2FF", "3730A3", "343C63"

# Page geometry in mm, as themes/aurora.css lays out an A4 sheet.
HEAD1, HEAD2, SIDE_W, BAR = (
    46,
    20,
    56,
    16,
)  # hero rows: name block, then headline numbers; sidebar width; bar height /10
HERO = 64  # the hero shape; the table rows end 1.5 mm short of it, Word starts the table that far down


def styles(doc: Doc) -> None:
    sans = partial(style, doc, base="Normal", font=SANS, color=INK, line=1.55, line_gap=SANS_GAP)
    serif = partial(sans, font=SERIF, line_gap=SERIF_GAP)

    _, dot, ring = bullet_lists(  # word.py's bullets take the first
        doc,
        [
            Mark("•", VIOLET, 624, 284, 10, "Segoe UI Symbol"),
            Mark("●", VIOLET, 300, 300, 7, "Segoe UI Symbol"),
            Mark("○", VIOLET, 340, 340, 12, "Segoe UI Symbol"),
        ],
    )

    sans("Normal", base=None, size=9.5, after=7)
    sans(
        "Name",
        base="Normal",
        font=SERIF,
        line_gap=SERIF_GAP,
        size=30,
        color="FFFFFF",
        bold=True,
        after=0,
        line=1.05,
        spacing=-0.6,
    )
    sans("Job Title", size=7.6, color=LAVENDER, caps=True, spacing=1.2, after=4, line=1.3)
    sans("CV Photo", size=2, after=0, line=1.0)
    contact = sans("CV Contact", size=8, color="E0E7FF", bold=True, before=0, after=0, line=1.3)
    contact.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    border(contact, "top between bottom", RULE, 6, 5)
    sans("CV Highlights", size=2, after=0)
    serif("CV Highlight Value", size=18, color="FFFFFF", bold=True, after=0, line=1.1)
    sans("CV Highlight Label", size=7.4, color=LAVENDER, after=0, line=1.3)

    numbered(serif("Heading 1", size=14, bold=True, before=0, after=10, line=1.2, spacing=-0.14), dot)
    sans("Heading 2", size=10.5, bold=True, before=4, after=0, line=1.3)
    numbered(sans("CV Role", base="Heading 2", size=10.5, bold=True, before=8, after=0, line=1.3), ring)
    sans("CV Company", size=8.3, color=MUTED, after=1, line=1.3).paragraph_format.left_indent = Cm(0.6)
    sans("Date", size=7.2, color=MUTED, bold=True, caps=True, spacing=0.3, after=6, line=1.3)
    role_date = sans(
        "CV Role Date", base="Date", size=7.2, color=MUTED, bold=True, caps=True, spacing=0.3, after=6, line=1.3
    )
    role_date.paragraph_format.left_indent = Cm(0.6)
    sans("Heading 3", size=6.8, color=MUTED, bold=True, caps=True, spacing=1.0, before=0, after=8, line=1.2)
    sans("Heading 4", size=6.6, color=MUTED, bold=True, caps=True, spacing=0.8, before=6, after=1, line=1.2)
    sans("Sidebar Text", size=8.7, after=1)
    kv = sans("CV Key Value", size=8.7, after=4)
    kv.paragraph_format.tab_stops.add_tab_stop(Cm(3.6), alignment=WD_TAB_ALIGNMENT.RIGHT)  # at the sidebar text edge
    sans("CV Skill", size=8.7, after=1)
    sans("CV Skill Bar", size=1, after=4, line=1.0)
    sans("List Paragraph", size=9.5, after=2).paragraph_format.left_indent = Cm(0.6)
    sans("CV Tools", size=7.2, after=10, line=1.8).paragraph_format.left_indent = Cm(0.6)
    quote = sans("CV Quote", size=9.5, color="3A3A3A", after=4)
    quote.font.italic = True
    quote.paragraph_format.left_indent = Cm(0.4)
    border(quote, "left", VIOLET, 18, 8)
    sans("Sidebar Sections Spacing", size=2, after=9)
    sans("Main Sections Spacing", size=4, after=12)
    char_style(doc, "CV Accent", font=SANS, color=TEAL)
    char_style(doc, "CV Muted", font=SANS, color=MUTED, size=7.5)
    char_style(doc, "CV Tool", font=SANS, fill=CHIP, size=7.2)
    char_style(doc, "CV Tool Mark", font=SANS, fill=VIOLET, size=7.2)  # a chip's violet left edge
    char_style(doc, "CV Skill Fill", font=SANS, fill=VIOLET)
    char_style(doc, "CV Skill Track", font=SANS, fill=LINE)
    for s in ("CV Company", "CV Role Date", "List Paragraph", "CV Tools"):
        border(doc.styles[s], "left", LINE, 6, 8)  # the web's hairline; its ring sits on it
    for s in ("CV Contact", "CV Role", "CV Company", "Date", "CV Role Date", "Heading 1", "Heading 3"):
        doc.styles[s].paragraph_format.keep_with_next = True


def header(doc: Doc) -> None:
    """Hero rows of fixed height, so the body starts right under the headline numbers."""
    head = table(doc, [14.0, 7.0], rows=2)
    for row, mm in zip(head.rows, (HEAD1, HEAD2), strict=True):
        row.height, row.height_rule = Cm(mm / 10), WD_ROW_HEIGHT_RULE.EXACTLY
    name_c, contact_c = head.rows[0].cells
    cell(name_c, 14.0, pad=(1.25, 0.4, 0, 1.4), valign=WD_CELL_VERTICAL_ALIGNMENT.CENTER)
    cell(contact_c, 7.0, pad=(0.95, 1.4, 0, 0.4), valign=WD_CELL_VERTICAL_ALIGNMENT.CENTER)
    name_c.paragraphs[0]._p.style = "JobTitle"
    name_c.paragraphs[0].add_run("{Job Title}")
    name_c.add_paragraph(style="Name").add_run("{Name}")
    contact_c.paragraphs[0]._p.style = "CVContact"
    contact_c.paragraphs[0].add_run("{Contact}")
    stats = head.rows[1].cells[0].merge(head.rows[1].cells[1])
    cell(stats, 21, pad=(0.2, 1.4, 0, 1.4))
    stats.paragraphs[0]._p.style = "CVHighlights"
    stats.paragraphs[0].add_run("{Highlights}")


def layout(doc: Doc) -> None:
    """Body: sidebar column and main column."""
    doc.add_paragraph(style="CV Photo")  # keeps Word from merging the two tables
    side, main = table(doc, [SIDE_W / 10, 21 - SIDE_W / 10]).rows[0].cells
    cell(side, SIDE_W / 10, pad=(0.2, 0.6, 0, 1.4))
    cell(main, 21 - SIDE_W / 10, pad=(0.65, 1.4, 0, 0.8))
    side.paragraphs[0]._p.style = "CVPhoto"  # word.py adds a section spacer of its own
    doc.add_paragraph(style="CV Photo")  # Word needs a paragraph between the last table and the section end


def decorate(doc: Doc) -> None:
    """Page decoration as shapes behind the text, repeated by the page headers. Page 1: the night hero and its
    two bars. Later pages: the top bars, and the body starts 14 mm down, as the web's later sheets do."""
    sec = doc.sections[0]
    first = sec.first_page_header.paragraphs[0]
    first._p.style = "CVPhoto"
    bars = lambda y, h: [(0, y, 140, h, VIOLET), (140, y, 70, h, TEAL)]  # noqa: E731
    for i, box in enumerate([(0, 0, 210, HERO, NIGHT), *bars(HERO - BAR / 10, BAR / 10)]):
        rect(first, box, 1001 + i)
    later = sec.header.paragraphs[0]
    later._p.style = "CVPhoto"
    later.paragraph_format.space_before = Cm(0.55)
    for i, box in enumerate(bars(0, 2)):
        rect(later, box, 1101 + i)


def make(out: Path) -> None:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(0)
    sec.top_margin, sec.bottom_margin = Cm(0), Cm(0.8)
    sec.header_distance, sec.footer_distance = Cm(0), Cm(0)
    sec.different_first_page_header_footer = True
    styles(doc)
    header(doc)
    layout(doc)
    decorate(doc)
    finish(doc, out, SANS, SERIF)
