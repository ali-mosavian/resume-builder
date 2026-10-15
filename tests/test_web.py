import re
from pathlib import Path

import pytest

from resume import web
from support import CV
from resume import data
from support import html


@pytest.mark.parametrize("theme", web.themes())
def test_page_content(theme: str) -> None:
    page = web.render([data.load(CV, "en")], theme)
    assert "<strong>demand forecasting platform</strong>" in page
    assert "**" not in page
    assert "reading &amp; baking" in page, "YAML text is not HTML-escaped"
    # Brief roles are not contiguous (a full role sits between them): the caption must appear once.
    assert page.count("data-earlier=") == 1, "earlier-roles caption repeated or missing"


def test_no_dashes() -> None:
    """Dates were printed "JUN 2024 <em dash> NU": the CV uses "-", never em or en dashes."""
    found = re.search(".{20}[–—].{20}", web.render([data.load(CV, lang) for lang in ("en", "de")], "bracket"))
    assert not found, found[0]


def test_two_languages(tmp_path: Path) -> None:
    """One page, both languages: the second starts hidden; the theme script and the language switch are inlined."""
    page = html(tmp_path / "cv.html", "aurora-v3", ["en", "de"]).read_text(encoding="utf-8")
    assert re.search(r'<div class="page" lang="de" data-ui="[^"]*" hidden>', page), "second language not hidden"
    assert re.search(r'<div class="page" lang="en" data-ui="[^"]*">', page), "first language hidden"
    assert page.count("<script>") == len(web.scripts("aurora-v3", multilingual=True))
    assert 'data-aka="GCP"' in page


def test_download_links() -> None:
    page = web.render([data.load(CV, lang) for lang in ("en", "de")], "bracket")
    assert 'href="katrin_vogel-resume-de.docx" download' in page
    assert 'onclick="print()"' not in page
