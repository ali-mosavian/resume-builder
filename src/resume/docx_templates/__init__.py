"""Template generators: each module's make() writes templates/<module>.docx with the style ids word.py expects.

Shared here: styles, bullet lists, borders, cells and tables as the generators build them.
"""

import uuid
import pkgutil
from typing import Any
from typing import cast
from pathlib import Path
from typing import NamedTuple
from collections.abc import Sequence

from docx.shared import Cm
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.table import Table
from docx.table import _Cell
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.shared import RGBColor
from docx.oxml import OxmlElement
from docx.document import Document as Doc
from docx.enum.style import WD_STYLE_TYPE
from docx.text.paragraph import Paragraph
from docx.opc.oxml import serialize_part_xml
from docx.styles.style import CharacterStyle
from docx.styles.style import ParagraphStyle
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from resume import ROOT
from resume.ooxml import CM
from resume.ooxml import xml
from resume.ooxml import shading


def names() -> list[str]:
    """Template names, as their themes are named: module aurora_v2 makes aurora-v2.docx."""
    return [m.name.replace("_", "-") for m in pkgutil.iter_modules(__path__)]


def style(
    doc: Doc,
    name: str,
    base: str | None = None,
    *,
    size: float,
    font: str,
    color: str,
    bold: bool = False,
    caps: bool = False,
    spacing: float = 0,
    before: float = 0,
    after: float = 2,
    line: float = 1.15,
    line_gap: float = 1.0,
) -> ParagraphStyle:
    """line: a multiple of the font's line height, or of its size given the font's line_gap (as CSS line-height)."""
    styles = doc.styles
    known = name in [x.name for x in styles]
    s = cast(ParagraphStyle, styles[name] if known else styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH))
    if base:
        s.base_style = styles[base]
    f = s.font
    f.name, f.size, f.bold, f.all_caps, f.italic = font, Pt(size), bold, caps, False
    f.color.rgb = RGBColor.from_string(color)
    rpr = s.element.get_or_add_rPr()
    fonts = rpr.get_or_add_rFonts()
    fonts.set(qn("w:eastAsia"), font)
    for theme in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):  # a theme font beats the named one
        fonts.attrib.pop(qn(theme), None)
    if spacing:
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(int(spacing * 20)))
        rpr.append(sp)
    pf = s.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(before), Pt(after), line / line_gap
    pf.keep_with_next = name.startswith("Heading") or name == "Date"
    return s


def char_style(
    doc: Doc, name: str, *, font: str, color: str | None = None, fill: str | None = None, size: float | None = None
) -> CharacterStyle:
    s = cast(CharacterStyle, doc.styles.add_style(name, WD_STYLE_TYPE.CHARACTER))
    s.font.name = font
    if color:
        s.font.color.rgb = RGBColor.from_string(color)
    if size:
        s.font.size = Pt(size)
    if fill:
        s.element.get_or_add_rPr().append(shading(fill))
    return s


class Mark(NamedTuple):
    """A bullet list's mark; left and hanging indents in twips, size in points."""

    glyph: str
    color: str
    left: int
    hanging: int
    size: float | None = None
    font: str | None = None


def bullet_lists(doc: Doc, marks: Sequence[Mark], first_id: int = 90) -> list[str]:
    """One list per mark, ahead of python-docx's own: word.py takes the first bullet list for its bullets."""
    numbering = doc.part.numbering_part.element
    ids = []
    for i, m in enumerate(marks):
        n = first_id + i
        font = f'<w:rFonts w:ascii="{m.font}" w:hAnsi="{m.font}"/>' if m.font else ""
        size = f'<w:sz w:val="{int(m.size * 2)}"/>' if m.size else ""
        numbering.insert(
            i,
            xml(
                f'<w:abstractNum w:abstractNumId="{n}"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
                f'<w:numFmt w:val="bullet"/><w:lvlText w:val="{m.glyph}"/><w:lvlJc w:val="left"/>'
                f'<w:pPr><w:ind w:left="{m.left}" w:hanging="{m.hanging}"/></w:pPr>'
                f'<w:rPr>{font}<w:color w:val="{m.color}"/>{size}</w:rPr></w:lvl></w:abstractNum>'
            ),
        )
        numbering.append(xml(f'<w:num w:numId="{n}"><w:abstractNumId w:val="{n}"/></w:num>'))
        ids.append(str(n))
    return ids


# Schema order of w:pPr children that may follow the ones added here; Word rejects out-of-order elements.
AFTER_NUMPR = (
    "w:suppressLineNumbers",
    "w:pBdr",
    "w:shd",
    "w:tabs",
    "w:suppressAutoHyphens",
    "w:kinsoku",
    "w:wordWrap",
    "w:overflowPunct",
    "w:topLinePunct",
    "w:autoSpaceDE",
    "w:autoSpaceDN",
    "w:bidi",
    "w:adjustRightInd",
    "w:snapToGrid",
    "w:spacing",
    "w:ind",
    "w:contextualSpacing",
    "w:mirrorIndents",
    "w:suppressOverlap",
    "w:jc",
    "w:textDirection",
    "w:textAlignment",
    "w:textboxTightWrap",
    "w:outlineLvl",
    "w:divId",
    "w:cnfStyle",
    "w:rPr",
    "w:sectPr",
    "w:pPrChange",
)


def ppr(s: ParagraphStyle, snippet: str, after: tuple[str, ...]) -> None:
    """Add a w:pPr child before its schema successors."""
    s.element.get_or_add_pPr().insert_element_before(xml(snippet), *after)


def numbered(s: ParagraphStyle, num_id: str) -> None:
    ppr(s, f'<w:numPr><w:numId w:val="{num_id}"/></w:numPr>', AFTER_NUMPR)


def border(s: ParagraphStyle, sides: str, color: str, size: int, space: int = 4) -> None:
    """sides: space-separated w:pBdr sides. Adjacent paragraphs with equal borders share one box, so ruled
    lists need "top between bottom"."""
    order = ("top", "left", "bottom", "right", "between")  # schema order
    edges = "".join(
        f'<w:{side} w:val="single" w:sz="{size}" w:space="{space}" w:color="{color}"/>'
        for side in sorted(sides.split(), key=order.index)
    )
    ppr(s, f"<w:pBdr>{edges}</w:pBdr>", AFTER_NUMPR[2:])


def cell(
    c: _Cell,
    width_cm: float,
    *,
    fill: str | None = None,
    pad: tuple[float, float, float, float] = (0, 0, 0, 0),
    borders: dict[str, tuple[str, int]] | None = None,
    valign: WD_CELL_VERTICAL_ALIGNMENT | None = None,
) -> None:
    """pad: top, right, bottom, left in cm; borders: side -> (color, eighths of a point)."""
    c.width = Cm(width_cm)
    tc_pr = c._tc.get_or_add_tcPr()
    order = ("top", "left", "bottom", "right")  # schema order
    if borders:
        tc_pr.append(
            xml(
                "<w:tcBorders>"
                + "".join(
                    f'<w:{side} w:val="single" w:sz="{borders[side][1]}" w:space="0" w:color="{borders[side][0]}"/>'
                    for side in order
                    if side in borders
                )
                + "</w:tcBorders>"
            )
        )
    if fill:
        tc_pr.append(shading(fill))
    margin = dict(zip(("top", "right", "bottom", "left"), pad, strict=True))
    tc_pr.append(
        xml(
            "<w:tcMar>"
            + "".join(f'<w:{side} w:w="{int(margin[side] * CM)}" w:type="dxa"/>' for side in order)
            + "</w:tcMar>"
        )
    )
    if valign:
        c.vertical_alignment = valign


def table(
    container: Doc, widths: Sequence[float], rows: int = 1, align: WD_TABLE_ALIGNMENT = WD_TABLE_ALIGNMENT.LEFT
) -> Table:
    """A fixed-layout table; cells pad themselves (see cell)."""
    t = container.add_table(rows=rows, cols=len(widths))
    t.alignment = align
    t.autofit = False
    # Word and LibreOffice size columns from the grid, not from cell widths alone.
    for col, w in zip(t._tbl.tblGrid.gridCol_lst, widths, strict=True):
        col.w = Cm(w)
    t._tbl.tblPr.get_or_add_tblLayout().set(qn("w:type"), "fixed")
    t._tbl.tblPr.insert_element_before(
        xml('<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'),
        "w:tblLook",
        "w:tblCaption",
        "w:tblDescription",
    )
    return t


FONTS = ROOT / "assets" / "fonts"
FACES = ("Regular", "Bold", "Italic", "BoldItalic")


def embed_font(doc: Doc, family: str) -> None:
    """Embed the family's faces from assets/fonts (<Family>-{Regular,Bold,Italic,BoldItalic}.ttf, see tools/fonts.py)
    so Word shows the font without it installed. Call once per family.

    Word stores each as an obfuscated .odttf: the first 32 bytes XORed with its GUID key, bytes reversed.
    """
    files = {f: p for f in FACES if (p := FONTS / f"{family.replace(' ', '')}-{f}.ttf").exists()}
    assert files, f"no faces of {family} in {FONTS}"
    font_table = doc.part.part_related_by(RT.FONT_TABLE)
    root = parse_xml(font_table.blob)
    for old in root.findall(qn("w:font")):  # a font table isn't an oxml class: no w: prefix in xpath
        if old.get(qn("w:name")) == family:
            root.remove(old)
    font = xml(
        f'<w:font w:name="{family}"><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>'
    )
    for kind, path in files.items():
        guid = str(uuid.uuid4()).upper()
        key = bytes.fromhex(guid.replace("-", ""))[::-1]
        data = bytearray(path.read_bytes())
        for j in range(32):
            data[j] ^= key[j % 16]
        part = Part(
            doc.part.package.next_partname("/word/fonts/font%d.odttf"),
            "application/vnd.openxmlformats-officedocument.obfuscatedFont",
            bytes(data),
            doc.part.package,
        )
        font.append(
            xml(f'<w:embed{kind} r:id="{font_table.relate_to(part, RT.FONT)}" w:fontKey="{{{guid}}}"/>', "w", "r")
        )
    root.append(font)
    font_table._blob = serialize_part_xml(root)
    settings = doc.settings.element
    if not settings.xpath("./w:embedTrueTypeFonts"):
        settings.xpath("./w:zoom")[0].addnext(xml("<w:embedTrueTypeFonts/>"))  # schema: right after zoom


WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


type Box = tuple[float, float, float, float, str]  # x, y, width, height in mm; colour


def rect(p: Paragraph, box: Box, shape_id: int) -> None:
    """A filled rectangle behind the text, placed on the page: what CSS draws with positioned boxes."""
    x, y, w, h, color = box

    def emu(mm: float) -> int:
        return int(mm * 36000)

    p._p.append(
        xml(
            f'<w:r xmlns:wps="{WPS}"><w:drawing>'
            f'<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{shape_id}" behindDoc="1" '
            f'locked="1" layoutInCell="0" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{emu(x)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{emu(y)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{emu(w)}" cy="{emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="{shape_id}" name="Shape {shape_id}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="{WPS}"><wps:wsp><wps:cNvSpPr/><wps:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>'
            f"</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>",
            "w",
            "wp",
            "a",
        )
    )


def bold_font(family: str) -> str | None:
    """The family's bold faces as a family of their own (tools/fonts.py), if there are any."""
    return f"{family} Bold" if (FONTS / f"{family.replace(' ', '')}Bold-Regular.ttf").exists() else None


def real_bold(doc: Doc) -> set[str]:
    """Bold styles switch to the font's bold family, not bold: Word for Mac ignores embedded bold faces and fakes bold
    from the regular one. Adds CVStrong, which word.py uses for **bold** text. Returns the bold families used."""
    styles = doc.styles.element
    by_id = {s.get(qn("w:styleId")): s for s in styles.findall(qn("w:style"))}

    def font_of(s: Any) -> str | None:
        while s is not None:
            if found := s.xpath("./w:rPr/w:rFonts/@w:ascii"):
                return str(found[0])
            base = s.xpath("./w:basedOn/@w:val")
            s = by_id.get(base[0]) if base else None
        return None

    used = set()
    for s in by_id.values():
        b = s.find(f"{qn('w:rPr')}/{qn('w:b')}")
        if b is None or b.get(qn("w:val")) in ("0", "false"):
            continue
        family = font_of(s)
        if not family or not (bold := bold_font(family)):
            continue
        rpr = s.find(qn("w:rPr"))
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            rpr.insert(0, fonts)
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            fonts.set(qn(attr), bold)
        for tag in ("w:b", "w:bCs"):
            for el in rpr.findall(qn(tag)):
                rpr.remove(el)
        used.add(bold)
    # ponytail: CVStrong is in the body font's bold family; **bold** inside a heading font would switch to it too.
    if (body := font_of(by_id.get("Normal"))) and (bold := bold_font(body)):
        char_style(doc, "CV Strong", font=bold)
        used.add(bold)
    return used


def finish(doc: Doc, out: Path, *fonts: str) -> None:
    """Settings every template needs, real bold, the fonts embedded, then save."""
    settings = doc.settings.element
    settings.xpath("./w:zoom")[0].set(
        qn("w:percent"), "100"
    )  # python-docx's blank document omits this required attribute
    # Word 2013+ layout. In the blank document's Word 2010 mode, a table's first-cell text aligns with the margin,
    # so the cell padding pushes the table, the sidebar text and its marks off the page's left edge.
    settings.xpath("./w:compat/w:compatSetting[@w:name='compatibilityMode']")[0].set(qn("w:val"), "15")
    for family in [*fonts, *sorted(real_bold(doc) - set(fonts))]:
        embed_font(doc, family)
    doc.save(str(out))
