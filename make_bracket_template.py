# /// script
# requires-python = ">=3.12"
# dependencies = ["python-docx~=1.1"]
# ///
"""Generate templates/bracket.docx, the Word twin of themes/bracket.css.

Beyond the style ids every template has, it opts into build_cv.py's optional parts: Photo, Contact and
Highlights placeholders, skill bars, a role timeline, tool chips and quote bars.
"""
import uuid
from pathlib import Path

from docx import Document
from lxml import etree
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

from make_pro_template import style

FONT = "Montserrat"
INK, MUTED, ACCENT, SIDE, TRACK, LINE = "262626", "7A7A7A", "F6B92B", "F4F3EF", "DDDBD3", "D9D9D9"


LINE_GAP = 1.219  # Montserrat's line height per em: Word's "multiple" spacing scales it, CSS line-height doesn't


def font(*, line: float = 1.6, **kw):
    """A paragraph style in Montserrat; line is CSS line-height (a multiple of the font size)."""
    return style(font=FONT, line=line / LINE_GAP, **kw)


# Schema order of w:pPr children that may follow the ones added here; Word rejects out-of-order elements.
AFTER_NUMPR = ("w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
               "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
               "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
               "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId",
               "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")


def ppr(s, xml: str, after: tuple[str, ...]) -> None:
    """Add a w:pPr child, given as XML without namespace declarations, before its schema successors."""
    root = xml.index(">")
    s.element.get_or_add_pPr().insert_element_before(parse_xml(f"{xml[:root]} {nsdecls('w')}{xml[root:]}"), *after)


def border(s, sides: str, color: str, size: int, space: int = 4) -> None:
    """sides: space-separated w:pBdr sides. Adjacent paragraphs with equal borders share one box, so ruled
    lists need "top between bottom"."""
    order = ("top", "left", "bottom", "right", "between")  # schema order
    edges = "".join(f'<w:{side} w:val="single" w:sz="{size}" w:space="{space}" w:color="{color}"/>'
                    for side in sorted(sides.split(), key=order.index))
    ppr(s, f"<w:pBdr>{edges}</w:pBdr>", AFTER_NUMPR[2:])


def char_style(doc, name: str, *, color: str | None = None, fill: str | None = None, size: float | None = None, bold=None):
    s = doc.styles.add_style(name, WD_STYLE_TYPE.CHARACTER)
    s.font.name, s.font.bold = FONT, bold
    if color:
        s.font.color.rgb = RGBColor.from_string(color)
    if size:
        s.font.size = Pt(size)
    if fill:
        s.element.get_or_add_rPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))
    return s


def marks(doc) -> dict[str, str]:
    """Bullet list first (build_cv.py takes the first bullet numbering), then the diamond and ring marks."""
    numbering = doc.part.numbering_part.element
    # One list per style: Word links a list to a single style and drops the marks of any other.
    defs = {"list": ("•", ACCENT, 624, 284, 10), "diamond": ("◆", ACCENT, 400, 400, 11), "diamond side": ("◆", ACCENT, 400, 400, 11),
            "ring": ("○", INK, 340, 340, 12)}
    for i, (key, (glyph, color, left, hanging, size)) in enumerate(defs.items()):
        n = 90 + i
        numbering.insert(i, parse_xml(
            f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{n}"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
            f'<w:numFmt w:val="bullet"/><w:lvlText w:val="{glyph}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol"/><w:color w:val="{color}"/>'
            f'<w:sz w:val="{size * 2}"/></w:rPr></w:lvl></w:abstractNum>'))
        numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{n}"><w:abstractNumId w:val="{n}"/></w:num>'))
    return {key: str(90 + i) for i, key in enumerate(defs)}


def numbered(s, num_id: str) -> None:
    ppr(s, f'<w:numPr><w:numId w:val="{num_id}"/></w:numPr>', AFTER_NUMPR)


def cell(c, width_cm: float, *, fill: str | None = None, pad=(0, 0, 0, 0), borders: dict | None = None, valign=None) -> None:
    """pad: top, right, bottom, left in cm; borders: side -> (color, eighths of a point)."""
    c.width = Cm(width_cm)
    tc_pr = c._tc.get_or_add_tcPr()
    if borders:
        order = ("top", "left", "bottom", "right")  # schema order
        tc_pr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}>' + "".join(
            f'<w:{side} w:val="single" w:sz="{borders[side][1]}" w:space="0" w:color="{borders[side][0]}"/>'
            for side in order if side in borders) + "</w:tcBorders>"))
    if fill:
        tc_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))
    margin = dict(zip(("top", "right", "bottom", "left"), pad))
    tc_pr.append(parse_xml(f'<w:tcMar {nsdecls("w")}>' + "".join(
        f'<w:{side} w:w="{int(margin[side] * 567)}" w:type="dxa"/>' for side in ("top", "left", "bottom", "right"))  # schema order
        + "</w:tcMar>"))
    if valign:
        c.vertical_alignment = valign


def table(container, widths: list[float], rows: int = 1):
    t = container.add_table(rows=rows, cols=len(widths))
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    for col, w in zip(t._tbl.tblGrid.gridCol_lst, widths):
        col.w = Cm(w)
    t._tbl.tblPr.get_or_add_tblLayout().set(qn("w:type"), "fixed")
    t._tbl.tblPr.insert_element_before(parse_xml(
        f'<w:tblCellMar {nsdecls("w")}><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'),
        "w:tblLook", "w:tblCaption", "w:tblDescription")
    return t


def embed_font(doc, family: str, files: dict[str, Path]) -> None:
    """Embed TrueType files (keys: Regular, Bold, Italic, BoldItalic) so Word shows the font without it installed.

    Word stores each as an obfuscated .odttf: the first 32 bytes XORed with its GUID key, bytes reversed.
    """
    table = doc.part.part_related_by(RT.FONT_TABLE)
    root = parse_xml(table.blob)
    for old in root.findall(qn("w:font")):
        if old.get(qn("w:name")) == family:
            root.remove(old)
    font = parse_xml(f'<w:font {nsdecls("w")} w:name="{family}"><w:charset w:val="00"/><w:family w:val="swiss"/>'
                     '<w:pitch w:val="variable"/></w:font>')
    for i, (kind, path) in enumerate(files.items(), 1):
        guid = str(uuid.uuid4()).upper()
        key = bytes.fromhex(guid.replace("-", ""))[::-1]
        data = bytearray(path.read_bytes())
        for j in range(32):
            data[j] ^= key[j % 16]
        part = Part(PackURI(f"/word/fonts/font{i}.odttf"), "application/vnd.openxmlformats-officedocument.obfuscatedFont",
                    bytes(data), doc.part.package)
        el = parse_xml(f'<w:embed{kind} {nsdecls("w", "r")} r:id="{table.relate_to(part, RT.FONT)}" w:fontKey="{{{guid}}}"/>')
        font.append(el)
    root.append(font)
    table._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    settings = doc.settings.element
    settings.find(qn("w:zoom")).addnext(parse_xml(f'<w:embedTrueTypeFonts {nsdecls("w")}/>'))  # schema: right after zoom


WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


def rect(p, x: float, y: float, w: float, h: float, color: str, shape_id: int) -> None:
    """A filled rectangle behind the text, placed on the page in mm: what CSS draws with positioned boxes."""
    emu = lambda mm: int(mm * 36000)  # noqa: E731
    p._p.append(parse_xml(
        f'<w:r {nsdecls("w", "wp", "a")} xmlns:wps="{WPS}"><w:drawing>'
        f'<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{shape_id}" behindDoc="1" '
        f'locked="1" layoutInCell="0" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{emu(x)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{emu(y)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{emu(w)}" cy="{emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{shape_id}" name="Shape {shape_id}"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic><a:graphicData uri="{WPS}"><wps:wsp><wps:cNvSpPr/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>'
        f'</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'))


# Page geometry in mm, as themes/bracket.css lays out an A4 sheet.
TOP, HEAD1, HEAD2, SIDE_W = 7, 45, 24, 64  # header rows: name block, then headline numbers; sidebar width
SIDE_Y = TOP + HEAD1 + HEAD2  # where the sidebar starts on page 1


def main() -> None:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(0)
    sec.top_margin, sec.bottom_margin = Cm(TOP / 10), Cm(1.2)
    sec.header_distance, sec.footer_distance = Cm(0), Cm(0)
    sec.different_first_page_header_footer = True
    ids = marks(doc)

    font(doc=doc, name="Normal", size=9.5, after=4, line=1.6)
    font(doc=doc, name="Name", base="Normal", size=27.5, bold=True, caps=True, after=0, line=1.0)
    font(doc=doc, name="Job Title", base="Normal", size=6.8, color=MUTED, caps=True, spacing=1.3, after=3)
    font(doc=doc, name="CV Photo", base="Normal", size=2, after=0, line=1.0)
    contact = font(doc=doc, name="CV Contact", base="Normal", size=8, bold=True, before=0, after=0, line=1.0)
    contact.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    border(contact, "top between bottom", INK, 6, 5)
    font(doc=doc, name="CV Highlights", base="Normal", size=2, after=0)
    value = font(doc=doc, name="CV Highlight Value", base="Normal", size=18, bold=True, after=4, line=1.0)
    value.font.underline = WD_UNDERLINE.THICK
    value.element.rPr.find(qn("w:u")).set(qn("w:color"), ACCENT)
    font(doc=doc, name="CV Highlight Label", base="Normal", size=7.4, color=MUTED, after=0)

    h1 = font(doc=doc, name="Heading 1", base="Normal", size=12.8, bold=True, before=0, after=8, line=1.2)
    numbered(h1, ids["diamond"])
    font(doc=doc, name="Heading 2", base="Normal", size=9.7, bold=True, before=4, after=0)
    role = font(doc=doc, name="CV Role", base="Heading 2", size=9.7, bold=True, before=8, after=0)
    numbered(role, ids["ring"])
    company = font(doc=doc, name="CV Company", base="Normal", size=8.2, color=MUTED, after=1)
    company.paragraph_format.left_indent = Cm(0.6)
    date = font(doc=doc, name="Date", base="Normal", size=7.2, bold=True, spacing=0.3, after=6)
    date.font.underline = WD_UNDERLINE.THICK
    date.element.rPr.find(qn("w:u")).set(qn("w:color"), ACCENT)
    role_date = font(doc=doc, name="CV Role Date", base="Date", size=7.2, bold=True, spacing=0.3, after=6)
    role_date.font.underline = WD_UNDERLINE.THICK
    role_date.element.rPr.find(qn("w:u")).set(qn("w:color"), ACCENT)
    role_date.paragraph_format.left_indent = Cm(0.6)  # under the role title, past the ring
    h3 = font(doc=doc, name="Heading 3", base="Normal", size=10.6, bold=True, before=0, after=6, line=1.2)
    numbered(h3, ids["diamond side"])
    font(doc=doc, name="Heading 4", base="Normal", size=6.6, color=MUTED, bold=True, caps=True, spacing=1.1, before=6, after=0)
    font(doc=doc, name="Sidebar Text", base="Normal", size=8.7, after=1)
    kv = font(doc=doc, name="CV Key Value", base="Normal", size=8.7, after=3)
    kv.paragraph_format.tab_stops.add_tab_stop(Cm(4.6), alignment=2)  # right, at the sidebar text edge
    font(doc=doc, name="CV Skill", base="Normal", size=8.7, after=2)
    font(doc=doc, name="CV Skill Bar", base="Normal", size=1, after=6, line=1.0)
    lst = font(doc=doc, name="List Paragraph", base="Normal", size=9.5, after=2)
    lst.paragraph_format.left_indent = Cm(0.6)
    tools = font(doc=doc, name="CV Tools", base="Normal", size=7.2, after=10, line=1.8)
    tools.paragraph_format.left_indent = Cm(0.6)
    quote = font(doc=doc, name="CV Quote", base="Normal", size=9.5, color="3A3A3A", after=3)
    quote.font.italic = True
    quote.paragraph_format.left_indent = Cm(0.4)
    border(quote, "left", ACCENT, 18, 8)
    font(doc=doc, name="Sidebar Sections Spacing", base="Normal", size=4, after=16)
    font(doc=doc, name="Main Sections Spacing", base="Normal", size=4, after=26)
    char_style(doc, "CV Accent", color=ACCENT)
    char_style(doc, "CV Muted", color=MUTED, size=7.5)
    char_style(doc, "CV Tool", fill=SIDE, size=7.2)
    char_style(doc, "CV Tool Mark", fill=ACCENT, size=7.2)  # a chip's yellow left edge
    char_style(doc, "CV Skill Fill", fill=ACCENT)
    char_style(doc, "CV Skill Track", fill=TRACK)
    for s in ("CV Contact", "CV Role", "CV Company", "Date", "CV Role Date", "Heading 1", "Heading 3"):
        doc.styles[s].paragraph_format.keep_with_next = True

    # Header rows of fixed height, so page 1's sidebar shape starts right under the headline numbers.
    head = table(doc, [5.7, 9.3, 6.0], rows=2)
    for row, mm in zip(head.rows, (HEAD1, HEAD2)):
        row.height, row.height_rule = Cm(mm / 10), WD_ROW_HEIGHT_RULE.EXACTLY
    photo_c, name_c, contact_c = head.rows[0].cells
    cell(photo_c, 5.7, pad=(0.9, 0, 0, 1.6))
    cell(name_c, 9.3, pad=(0.9, 0.4, 0, 0.6), valign=WD_CELL_VERTICAL_ALIGNMENT.CENTER)
    cell(contact_c, 6.0, pad=(0.9, 1.4, 0, 0.4), valign=WD_CELL_VERTICAL_ALIGNMENT.CENTER)
    photo_c.paragraphs[0]._p.style = "CVPhoto"
    photo_c.paragraphs[0].add_run("{Photo 3.1 gray}")
    name_c.paragraphs[0]._p.style = "JobTitle"
    name_c.paragraphs[0].add_run("{Job Title}")
    name_c.add_paragraph(style="Name").add_run("{Name}")
    contact_c.paragraphs[0]._p.style = "CVContact"
    contact_c.paragraphs[0].add_run("{Contact}")
    stats = head.rows[1].cells[0].merge(head.rows[1].cells[2])
    cell(stats, 21, pad=(0.4, 1.4, 0, 1.6))
    stats.paragraphs[0]._p.style = "CVHighlights"
    stats.paragraphs[0].add_run("{Highlights}")

    # Body: sidebar column (its gray is a shape, see below) and main column.
    doc.add_paragraph(style="CV Photo")  # keeps Word from merging the two tables
    lay = table(doc, [SIDE_W / 10, 21 - SIDE_W / 10])
    side, main_c = lay.rows[0].cells
    cell(side, SIDE_W / 10, pad=(0.3, 0.8, 0, 1.0))
    cell(main_c, 21 - SIDE_W / 10, pad=(1.0, 1.4, 0, 1.0))
    side.paragraphs[0]._p.style = "CVPhoto"  # build_cv adds a section spacer of its own
    doc.add_paragraph(style="CV Photo")  # Word needs a paragraph between the last table and the section end

    # Page decoration as shapes behind the text, repeated by the page headers. Page 1: the bracket, its
    # black square, the photo's offset block, the sidebar from under the header down. Later pages: the
    # sidebar from the top, and the body starts 14 mm down, as the web's later sheets do.
    first = sec.first_page_header.paragraphs[0]
    first._p.style = "CVPhoto"
    for i, (x, y, w, h, color) in enumerate([(7, 7, 60, 5, ACCENT), (7, 7, 5, 30, ACCENT), (14, 12, 2.2, 2.2, INK),
                                              (19, 19, 31, 31, ACCENT), (0, SIDE_Y, SIDE_W, 297 - SIDE_Y, SIDE),
                                              (188, 292, 22, 5, ACCENT)]):
        rect(first, x, y, w, h, color, 1001 + i)
    later = sec.header.paragraphs[0]
    later._p.style = "CVPhoto"
    later.paragraph_format.space_before = Cm(1.3)
    for i, (x, y, w, h, color) in enumerate([(0, 0, SIDE_W, 297, SIDE), (188, 292, 22, 5, ACCENT)]):
        rect(later, x, y, w, h, color, 1101 + i)
    zoom = doc.settings.element.find(qn("w:zoom"))
    zoom.set(qn("w:percent"), "100")  # python-docx's blank document omits this required attribute
    # Word 2013+ layout. In the blank document's Word 2010 mode, a table's first-cell text aligns with the margin,
    # so the cell padding pushes the table, the sidebar text and its marks off the page's left edge.
    doc.settings.element.xpath("./w:compat/w:compatSetting[@w:name='compatibilityMode']")[0].set(qn("w:val"), "15")
    fonts = Path(__file__).parent / "assets" / "fonts"
    embed_font(doc, FONT, {kind: fonts / f"{FONT}-{kind}.ttf" for kind in ("Regular", "Bold", "Italic", "BoldItalic")})
    doc.save(Path(__file__).parent / "templates" / "bracket.docx")


if __name__ == "__main__":
    main()
