"""Fetch the themes' Google fonts and write the four static faces Word can embed: <Family>-{Regular,Bold,Italic,BoldItalic}.ttf.

Variable fonts are pinned at wght 400/700 and every other axis at its default, which is what the web pages use with
`font-optical-sizing: none`. Faces a family lacks (Instrument Serif has no bold) are skipped; Word synthesizes them.

Word for Mac ignores embedded bold and italic faces and fakes them from the regular one; only a family's regular face
is used. So each bold face is also written as the regular (and italic) face of its own family, "<Family> Bold", which
templates use for bold text. Faces already in assets/fonts are not fetched again.
Run: uv run python tools/fonts.py
"""

import io
import urllib.parse
import urllib.request
from typing import Any
from pathlib import Path

from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont

OUT = Path(__file__).resolve().parents[1] / "assets" / "fonts"
RAW = "https://raw.githubusercontent.com/google/fonts/main/ofl"

# family -> (directory, upright file, italic file); a None file means the family has no such face.
VARIABLE = {
    "Fraunces": ("fraunces", "Fraunces[SOFT,WONK,opsz,wght].ttf", "Fraunces-Italic[SOFT,WONK,opsz,wght].ttf"),
    "Inter": ("inter", "Inter[opsz,wght].ttf", "Inter-Italic[opsz,wght].ttf"),
    "Instrument Sans": ("instrumentsans", "InstrumentSans[wdth,wght].ttf", "InstrumentSans-Italic[wdth,wght].ttf"),
    "JetBrains Mono": ("jetbrainsmono", "JetBrainsMono[wght].ttf", "JetBrainsMono-Italic[wght].ttf"),
    "IBM Plex Sans": ("ibmplexsans", "IBMPlexSans[wdth,wght].ttf", "IBMPlexSans-Italic[wdth,wght].ttf"),
    "Newsreader": ("newsreader", "Newsreader[opsz,wght].ttf", "Newsreader-Italic[opsz,wght].ttf"),
    "Bricolage Grotesque": ("bricolagegrotesque", "BricolageGrotesque[opsz,wdth,wght].ttf", None),
    "Inter Tight": ("intertight", "InterTight[wght].ttf", "InterTight-Italic[wght].ttf"),
    "Space Grotesk": ("spacegrotesk", "SpaceGrotesk[wght].ttf", None),
}
# family -> (directory, {face: file})
STATIC = {
    "Instrument Serif": (
        "instrumentserif",
        {"Regular": "InstrumentSerif-Regular.ttf", "Italic": "InstrumentSerif-Italic.ttf"},
    ),
    "IBM Plex Mono": ("ibmplexmono", {f: f"IBMPlexMono-{f}.ttf" for f in ("Regular", "Bold", "Italic", "BoldItalic")}),
    "Lato": ("lato", {f: f"Lato-{f}.ttf" for f in ("Regular", "Bold", "Italic", "BoldItalic")}),
}
FACES = {"Regular": (400, False), "Bold": (700, False), "Italic": (400, True), "BoldItalic": (700, True)}


def fetch(directory: str, name: str) -> TTFont:
    url = f"{RAW}/{directory}/{urllib.parse.quote(name)}"
    with urllib.request.urlopen(url) as r:
        return TTFont(io.BytesIO(r.read()))


def rename(font: TTFont, family: str, face: str) -> None:
    """Name the face as one of a four-member Word family, so Word's bold and italic buttons find it."""
    weight, italic = FACES[face]
    style = {"BoldItalic": "Bold Italic"}.get(face, face)
    postscript = f"{family.replace(' ', '')}-{face}"
    names = font["name"]
    for record in list(names.names):
        if record.nameID in (1, 2, 3, 4, 6, 16, 17, 21, 22, 25):
            names.removeNames(nameID=record.nameID)
    for name_id, value in ((1, family), (2, style), (3, postscript), (4, f"{family} {style}"), (6, postscript)):
        names.setName(value, name_id, 3, 1, 0x409)
        names.setName(value, name_id, 1, 0, 0)
    os2: Any = font["OS/2"]  # fontTools tables get their fields at decompile time
    head: Any = font["head"]
    os2.usWeightClass = weight
    os2.fsSelection = (os2.fsSelection & ~0b1100001) | (1 if italic else 0) | (0b100000 if weight == 700 else 0)
    if not italic and weight == 400:
        os2.fsSelection |= 0b1000000  # REGULAR
    head.macStyle = (2 if italic else 0) | (1 if weight == 700 else 0)
    if "STAT" in font:
        del font["STAT"]


def path(family: str, face: str) -> Path:
    return OUT / f"{family.replace(' ', '')}-{face}.ttf"


def save(font: TTFont, family: str, face: str) -> None:
    rename(font, family, face)
    font.save(path(family, face))
    print(f"wrote {path(family, face).name}")


def bold_family(family: str) -> None:
    """<Family> Bold: the bold faces as a family of their own, for Word for Mac."""
    for face, as_face in (("Bold", "Regular"), ("BoldItalic", "Italic")):
        if path(family, face).exists():
            save(TTFont(path(family, face)), f"{family} Bold", as_face)


def main() -> None:
    for family, (directory, upright, italic) in VARIABLE.items():
        for face, (weight, slanted) in FACES.items():
            name = italic if slanted else upright
            if name is None or path(family, face).exists():
                continue
            font = fetch(directory, name)
            axes = {a.axisTag: a.defaultValue for a in font["fvar"].axes} | {"wght": weight}
            save(instantiateVariableFont(font, axes), family, face)
    for family, (directory, files) in STATIC.items():
        for face, name in files.items():
            if not path(family, face).exists():
                save(fetch(directory, name), family, face)
    for family in [*VARIABLE, *STATIC, "Montserrat"]:
        bold_family(family)


if __name__ == "__main__":
    main()
