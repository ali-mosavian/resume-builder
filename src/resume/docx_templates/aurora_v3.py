"""templates/aurora-v3.docx, the Word twin of themes/aurora-v3.css.

A dark hero band (shapes behind the text), links and headline numbers in it, then sidebar and main columns.
Opts into word.py's optional parts: Contact and Highlights placeholders, skill bars, role timeline, tool chips, quote bars.
"""

from pathlib import Path
from functools import partial

from docx import Document
from docx.shared import Cm
from docx.document import Document as Doc
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

from resume.ooxml import xml
from resume.docx_templates import WPS
from resume.docx_templates import Mark
from resume.docx_templates import cell
from resume.docx_templates import style
from resume.docx_templates import table
from resume.docx_templates import border
from resume.docx_templates import finish
from resume.docx_templates import numbered
from resume.docx_templates import char_style
from resume.docx_templates import bullet_lists

SERIF, SANS = "Fraunces", "Inter"
INK, MUTED, VIOLET, TEAL, LINE = "151A2D", "667085", "7C5CFF", "19C3B1", "E4E7EF"
CHIP, CHIP_INK, QUOTE = "EEF2FF", "3730A3", "A5B4FC"
LIGHT = "A5B4FC"
INTER_GAP, FRAUNCES_GAP = 1.21, 1.233  # line height per em: Word's "multiple" scales it, CSS line-height doesn't

# Page geometry in mm, as themes/aurora-v3.css lays out an A4 sheet.
TOP, HEAD1, HEAD2, SIDE_W = 7, 33, 22, 56  # top margin; name block; headline numbers row; sidebar width
HERO = TOP + HEAD1 + HEAD2


def shape(p: Paragraph, box: tuple[float, float, float, float], fill: str, shape_id: int, prst: str = "rect") -> None:
    """A filled shape behind the text. fill: a DrawingML fill element. Shapes past the page edge are clipped."""
    x, y, w, h = (int(v * 36000) for v in box)
    p._p.append(
        xml(
            f'<w:r xmlns:wps="{WPS}"><w:drawing>'
            f'<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{shape_id}" behindDoc="1" '
            f'locked="1" layoutInCell="0" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{y}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{w}" cy="{h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="{shape_id}" name="Shape {shape_id}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="{WPS}"><wps:wsp><wps:cNvSpPr/><wps:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>'
            f"{fill}<a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>"
            f"</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>",
            "w",
            "wp",
            "a",
        )
    )


def solid(color: str) -> str:
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


def sweep(start: str, end: str) -> str:
    """A 135deg linear gradient, as CSS draws it."""
    stops = "".join(f'<a:gs pos="{pos}"><a:srgbClr val="{c}"/></a:gs>' for pos, c in ((0, start), (100000, end)))
    return f'<a:gradFill rotWithShape="1"><a:gsLst>{stops}</a:gsLst><a:lin ang="2700000" scaled="0"/></a:gradFill>'


def styles(doc: Doc) -> None:
    sans = partial(style, doc, base="Normal", font=SANS, color=INK, line=1.5, line_gap=INTER_GAP)
    serif = partial(style, doc, base="Normal", font=SERIF, color=INK, line_gap=FRAUNCES_GAP)

    # One list per style: Word links a list to a single style and drops the marks of any other.
    _, ring = bullet_lists(
        doc,
        [Mark("•", VIOLET, 624, 284, 10, "Segoe UI Symbol"), Mark("○", VIOLET, 340, 340, 12, "Segoe UI Symbol")],
    )

    sans("Normal", base=None, size=9, after=4)
    serif("Name", size=30, color="FFFFFF", bold=True, spacing=-0.6, after=0, line=1.0)
    sans("Job Title", size=7, color=LIGHT, bold=True, caps=True, spacing=1.1, after=3, line=1.4)
    sans("CV Photo", size=2, after=0, line=1.0)
    contact = sans("CV Contact", size=8, color="FFFFFF", bold=True, after=0, line=1.4)
    contact.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    border(contact, "top between bottom", "3A4163", 6, 5)
    sans("CV Highlights", size=2, after=0)
    serif("CV Highlight Value", size=17, color="FFFFFF", bold=True, after=0, line=1.1)
    sans("CV Highlight Label", size=7, color=LIGHT, after=0)

    serif("Heading 1", size=14, bold=True, before=0, after=8, line=1.2)
    serif("Heading 2", size=11, bold=True, before=4, after=0, line=1.3)
    role = serif("CV Role", base="Heading 2", size=10, bold=True, before=6, after=0, line=1.3)
    numbered(role, ring)
    sans("CV Company", size=9, color="344054", after=1).paragraph_format.left_indent = Cm(0.6)
    sans("Date", size=7, color=MUTED, bold=True, caps=True, spacing=0.5, before=1, after=5, line=1.3)
    sans(
        "CV Role Date", base="Date", size=7, color=MUTED, bold=True, caps=True, spacing=0.5, before=1, after=5, line=1.3
    )
    doc.styles["CV Role Date"].paragraph_format.left_indent = Cm(0.6)
    sans("Heading 3", size=6.5, color=MUTED, bold=True, caps=True, spacing=0.9, before=0, after=6, line=1.2)
    sans("Heading 4", size=6.5, color=MUTED, bold=True, caps=True, spacing=0.8, before=6, after=0)
    sans("Sidebar Text", size=8.5, after=1)
    kv = sans("CV Key Value", size=8.5, after=4)
    kv.paragraph_format.tab_stops.add_tab_stop(Cm(3.6), alignment=2)  # right-aligned, at the sidebar text edge
    sans("CV Skill", size=8.5, after=1)
    sans("CV Skill Bar", size=1, after=6, line=1.0)
    sans("List Paragraph", size=9, after=4).paragraph_format.left_indent = Cm(0.6)
    sans("CV Tools", size=7, after=10, line=1.9).paragraph_format.left_indent = Cm(0.6)
    quote = sans("CV Quote", size=9, color="374151", after=4)
    quote.font.italic = True
    quote.paragraph_format.left_indent = Cm(0.4)
    border(quote, "left", QUOTE, 18, 8)
    sans("Sidebar Sections Spacing", size=4, after=8)
    sans("Main Sections Spacing", size=4, after=16)
    char_style(doc, "CV Accent", font=SANS, color=TEAL)
    char_style(doc, "CV Muted", font=SANS, color=MUTED, size=7.5)
    char_style(doc, "CV Tool", font=SANS, color=CHIP_INK, fill=CHIP, size=7)
    char_style(doc, "CV Tool Mark", font=SANS, fill=CHIP, size=7)
    char_style(doc, "CV Skill Fill", font=SANS, fill=VIOLET)
    char_style(doc, "CV Skill Track", font=SANS, fill=LINE)
    for s in ("CV Contact", "CV Role", "CV Company", "Date", "CV Role Date", "Heading 1", "Heading 3"):
        doc.styles[s].paragraph_format.keep_with_next = True


def header(doc: Doc) -> None:
    """Hero rows of fixed height, so the body starts right under the band."""
    head = table(doc, [15.5, 5.5], rows=2)
    for row, mm in zip(head.rows, (HEAD1, HEAD2), strict=True):
        row.height, row.height_rule = Cm(mm / 10), WD_ROW_HEIGHT_RULE.EXACTLY
    name_c, contact_c = head.rows[0].cells
    cell(name_c, 15.5, pad=(0, 0.4, 0, 1.4), valign=WD_CELL_VERTICAL_ALIGNMENT.CENTER)
    cell(contact_c, 5.5, pad=(0, 1.4, 0, 0.2), valign=WD_CELL_VERTICAL_ALIGNMENT.CENTER)
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
    doc.add_paragraph(style="CV Photo")  # keeps Word from merging the two tables
    side, main = table(doc, [SIDE_W / 10, 21 - SIDE_W / 10]).rows[0].cells
    cell(side, SIDE_W / 10, pad=(0.8, 0.6, 0, 1.4))
    cell(main, 21 - SIDE_W / 10, pad=(0.8, 1.4, 0, 0.6))
    side.paragraphs[0]._p.style = "CVPhoto"  # word.py adds a section spacer of its own
    doc.add_paragraph(style="CV Photo")  # Word needs a paragraph between the last table and the section end


def decorate(doc: Doc) -> None:
    """The hero band and its glows, on page 1 only; later pages start 14 mm down, as the web's later sheets do."""
    sec = doc.sections[0]
    first = sec.first_page_header.paragraphs[0]
    first._p.style = "CVPhoto"
    strip_y = TOP + HEAD1 + 2
    for i, (box, fill, prst) in enumerate(
        [
            ((0, 0, 210, HERO), sweep("0B1024", "1A2250"), "rect"),
            ((14, strip_y, 182, 16), solid("10173A"), "rect"),
        ]
    ):
        shape(first, box, fill, 1001 + i, prst)
    later = sec.header.paragraphs[0]
    later._p.style = "CVPhoto"
    later.paragraph_format.space_before = Cm(0.7)


def make(out: Path) -> None:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(0)
    sec.top_margin, sec.bottom_margin = Cm(TOP / 10), Cm(1.2)
    sec.header_distance, sec.footer_distance = Cm(0), Cm(0)
    sec.different_first_page_header_footer = True
    styles(doc)
    header(doc)
    layout(doc)
    decorate(doc)
    finish(doc, out, SERIF, SANS)
