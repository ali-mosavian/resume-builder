"""The page in headless Chrome: printing, pagination into A4 sheets and the theme scripts."""

import re
import subprocess
from pathlib import Path

import pytest
from typer.testing import CliRunner

from support import CV
from resume import data
from support import ROOT
from support import html
from support import title
from resume.cli import app
from resume import chrome as browser

THEMES = ROOT / "themes"


def sharing(script: str) -> list[str]:
    """Themes with a script (<theme>.js or <theme>.*.js) that is a link to the given one."""
    return sorted({p.name.split(".")[0] for p in THEMES.glob("*.js") if p.resolve() == (THEMES / script).resolve()})


def pdf(chrome: Path, page: Path, budget: int = 4000) -> Path:
    out = page.with_suffix(".pdf")
    browser.pdf(chrome, page.as_uri(), out, budget)
    return out


def pdftotext(*args: str | Path) -> str:
    return subprocess.run(["pdftotext", *map(str, args), "-"], capture_output=True, text=True, check=True).stdout


def test_headline_numbers_print_in_full(chrome: Path, tmp_path: Path) -> None:
    """Printing during the headline count-up once put partial numbers ("2+", "7M") in the PDF."""
    printed = pdf(chrome, html(tmp_path / "cv.html", "aurora", ["en"]), budget=300)
    assert {"12+", "8", "40+"} <= set(pdftotext("-l", "1", printed).split())


@pytest.mark.parametrize("theme", sharing("slate.js"))
def test_paginated_print_keeps_everything(chrome: Path, tmp_path: Path, theme: str) -> None:
    """slate.js paginates into A4 sheets that clip overflow; chained testimonials once overflowed one sheet
    and the last four vanished from the PDF. Every role and testimonial must reach the paper, on A4."""
    printed = pdf(chrome, html(tmp_path / "cv.html", theme, ["en"]))
    # Raw (content-stream) order: the default reading order once wove sidebar lines into a wrapped role heading
    # and reported a role lost. Raw drops spaces, so compare without whitespace.
    text = re.sub(r"\s", "", pdftotext("-raw", printed))
    cv = data.load(CV, "en")
    names = [*(j["company"] for j in cv["experience"]), *(t["name"] for t in cv["testimonials"])]
    assert not [n for n in names if re.sub(r"\s", "", n) not in text]
    info = subprocess.run(["pdfinfo", printed], capture_output=True, text=True, check=True).stdout
    assert "(A4)" in info


def test_split_prints_one_column(chrome: Path, tmp_path: Path) -> None:
    """split is two halves on screen and one column on paper: every role and testimonial reaches the A4 PDF, the
    body comes before the sidebar's sections (as split.docx), and earlier roles print in full."""
    printed = pdf(chrome, html(tmp_path / "cv.html", "split", ["en"]))
    text = re.sub(r"\s", "", pdftotext("-raw", printed))
    cv = data.load(CV, "en")
    names = [*(j["company"] for j in cv["experience"]), *(t["name"] for t in cv["testimonials"])]
    assert not [n for n in names if re.sub(r"\s", "", n) not in text]
    assert text.index("PROFESSIONALSUMMARY") < text.index("AREASOFEXPERTISE")
    brief = next(j for j in cv["experience"] if j.get("brief"))
    assert re.sub(r"\s|\*", "", brief["bullets"][0])[:30] in text
    info = subprocess.run(["pdfinfo", printed], capture_output=True, text=True, check=True).stdout
    assert "(A4)" in info


def test_bracket_sidebar_on_every_sheet(chrome: Path, tmp_path: Path) -> None:
    """bracket once dropped the sidebar column on sheets after its content ran out."""
    probe = (
        "addEventListener('load',()=>setTimeout(()=>{document.title=[...document.querySelectorAll('.sheet aside')]"
        ".map(a=>a.getBoundingClientRect().width|0).join(',')},1500))"
    )
    widths = [int(w) for w in title(chrome, html(tmp_path / "cv.html", "bracket", ["en"], probe).as_uri()).split(",")]
    assert len(widths) > 1
    assert min(widths) > 100, widths


def test_language_switch_paginates(chrome: Path, tmp_path: Path) -> None:
    """The language opened by #de is cut into sheets, and so is the one switched to."""
    probe = (
        "addEventListener('load',()=>setTimeout(()=>{const n=()=>[...document.querySelectorAll('.page')]"
        ".map(p=>p.lang+(p.hidden?'-':'+')+p.querySelectorAll('.sheet').length).join(' ');const a=n();"
        "const x=!!document.querySelector('.page[lang=de] .sheet:first-child [data-from=experience]');"
        "document.querySelector('.page:not([hidden]) .lang-switch [data-lang=en]').click();"
        "setTimeout(()=>{const y=!!document.querySelector('.page[lang=en] .sheet:first-child [data-from=experience]');"
        "document.title=a+'|'+n()+'|'+x+'|'+y},1000)},1500))"
    )
    page = html(tmp_path / "cv.html", "bracket", ["en", "de"], probe)
    before, after, de_first, en_first = title(chrome, page.as_uri() + "#de", budget=6000).split("|")
    # Experience once jumped to sheet 2, leaving a third of sheet 1 empty: in Swedish (where it was found) because the role's bottom
    # padding was counted, in English because a long first bullet could not break across sheets.
    assert de_first == en_first == "true", f"Experience not on sheet 1: de {de_first}, en {en_first}"
    assert before.startswith("en-0 de+")
    assert not before.endswith("+0"), before
    assert after.startswith("en+")
    assert not after.startswith("en+0"), after


@pytest.mark.parametrize("theme", sharing("aurora.js"))
def test_expanded_brief_roles_keep_tool_layout(chrome: Path, tmp_path: Path, theme: str) -> None:
    """ "Show details" on earlier roles once turned their tool tags into a plain block (display:block beat flex)."""
    probe = (
        "addEventListener('load',()=>{const d=e=>getComputedStyle(e).display;"
        "document.querySelector('.page:not([hidden]) .details-toggle').click();"
        "document.title=d(document.querySelector('.item:not(.brief) .tools'))+'|'+d(document.querySelector('.brief .tools'))})"
    )
    normal, brief = title(chrome, html(tmp_path / "cv.html", theme, ["en"], probe).as_uri(), budget=1000).split("|")
    assert normal == brief


@pytest.mark.parametrize("theme", sorted(set(sharing("aurora.js")) & set(sharing("slate.js"))))
def test_skill_filter_after_pagination(chrome: Path, tmp_path: Path, theme: str) -> None:
    """Pagination into sheets replaced the roles the filter had captured: clicking a skill marked nothing on the
    page and the pill counted 0 roles."""
    probe = (
        "addEventListener('load',()=>setTimeout(()=>{const p=document.querySelector('.page:not([hidden])');"
        "[...p.querySelectorAll('aside .kv[style]')].find(k=>k.textContent.includes('Python')).click();"
        "document.title=p.querySelectorAll('.sheet .experience .item.match').length+'|'+p.querySelector('.filter-pill').textContent},1500))"
    )
    marked, pill = title(chrome, html(tmp_path / "cv.html", theme, ["en"], probe).as_uri()).split("|", 1)
    assert int(marked) > 0
    assert not pill.rstrip("Clear").rstrip().endswith("· 0"), pill


@pytest.mark.parametrize("theme", [p.stem for p in sorted(THEMES.glob("*.css"))])
def test_one_language_shown_with_a_switch(chrome: Path, tmp_path: Path, theme: str) -> None:
    """slate's `display:grid` beat `[hidden]`, so both languages showed (a 15-page PDF), and its header hid the
    language switch and download links."""
    probe = (
        "addEventListener('load',()=>setTimeout(()=>{const v=e=>!!e&&e.getClientRects().length>0;"
        "document.title=[...document.querySelectorAll('.page')].filter(v).length+'|'"
        "+v(document.querySelector('.page:not([hidden]) .lang-switch'))+'|'+v(document.querySelector('.page:not([hidden]) a.download'))},1500))"
    )
    pages, switch, download = title(chrome, html(tmp_path / "cv.html", theme, ["en", "de"], probe).as_uri()).split("|")
    assert (pages, switch, download) == ("1", "true", "true")


def test_build_writes_every_output(chrome: Path, tmp_path: Path) -> None:
    result = CliRunner().invoke(app, ["build", "--out", str(tmp_path)])
    assert result.exit_code == 0, result.output
    names = [data.load(CV, lang)["download"] for lang in data.langs(CV)]
    assert sorted(p.name for p in tmp_path.iterdir()) == sorted(
        ["index.html", *(f"{n}.{ext}" for n in names for ext in ("pdf", "docx"))]
    )
    assert all(p.stat().st_size for p in tmp_path.iterdir())
