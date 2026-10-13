"""Shared by the tests: the CV, builds of it, and headless Chrome."""

import re
from pathlib import Path

from docx.document import Document as Doc

from resume import ROOT, data, web, word
from resume.chrome import headless

CV = ROOT / "cv" / "cv.yaml"
TEMPLATES = ROOT / "templates"


def docx(template: str, lang: str = "en") -> Doc:
    return word.render(data.load(CV, lang), TEMPLATES / f"{template}.docx")


def html(out: Path, theme: str, langs: list[str], probe: str = "") -> Path:
    """The page, with an optional script that reports through document.title."""
    page = web.render([data.load(CV, lang) for lang in langs], theme)
    out.write_text(page.replace("</body>", f"<script>{probe}</script></body>") if probe else page, encoding="utf-8")
    return out


def title(chrome: Path, url: str, budget: int = 4000) -> str:
    """What a probe left in document.title."""
    found = re.search("<title>(.*?)</title>", headless(chrome, "--dump-dom", url, budget=budget))
    assert found, "page has no title"
    return found[1]
