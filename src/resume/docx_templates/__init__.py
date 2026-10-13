"""Template generators: each module's make() writes templates/<module>.docx with the style ids word.py expects.

Shared here: styles, bullet lists, borders, cells and tables as the generators build them.
"""

import pkgutil
from typing import cast
from typing import NamedTuple
from collections.abc import Sequence

from docx.shared import Cm
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.table import Table
from docx.table import _Cell
from docx.shared import RGBColor
from docx.oxml import OxmlElement
from docx.document import Document as Doc
from docx.enum.style import WD_STYLE_TYPE
from docx.styles.style import CharacterStyle
from docx.styles.style import ParagraphStyle
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

from resume.ooxml import CM
from resume.ooxml import xml
from resume.ooxml import shading


def names() -> list[str]:
    return [m.name for m in pkgutil.iter_modules(__path__)]


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
