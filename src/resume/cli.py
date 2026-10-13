"""resume: build the CV's web page, PDFs and .docx files, and the .docx templates."""

import importlib
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Annotated

import typer

from resume import ROOT, chrome, data, docx_templates, web, word

app = typer.Typer(add_completion=False, no_args_is_help=True)


def docx(cv: data.CV, template: Path, out: Path) -> None:
    word.render(cv, template).save(str(out))


@app.command()
def build(
    out: Annotated[Path, typer.Option(help="Output directory.")] = ROOT / "build",
    theme: Annotated[str, typer.Option(help=f"One of: {', '.join(web.themes())}")] = "bracket",
    template: Annotated[Path, typer.Option(help="A .docx made by `resume template`, or any with its style ids.")] = ROOT
    / "templates"
    / "bracket.docx",
    cv: Annotated[Path, typer.Option(help="The CV's YAML.")] = ROOT / "cv" / "cv.yaml",
) -> None:
    """index.html with every language, and per language a PDF of it and a .docx, named by the YAML's `download`."""
    browser = chrome.find()
    if browser is None:
        raise typer.BadParameter(f"no Chrome to print PDFs with; looked for {', '.join(chrome.CANDIDATES)}")
    out.mkdir(parents=True, exist_ok=True)
    cvs = [data.load(cv, lang) for lang in data.langs(cv)]
    page = out / "index.html"
    page.write_text(web.render(cvs, theme), encoding="utf-8")
    written = [page]
    with ThreadPoolExecutor() as pool:
        jobs = []
        for c in cvs:
            name = c.get("download", f"cv-{c['lang']}")
            pdf, doc = out / f"{name}.pdf", out / f"{name}.docx"
            jobs += [pool.submit(chrome.pdf, browser, f"{page.resolve().as_uri()}#{c['lang']}", pdf), pool.submit(docx, c, template, doc)]
            written += [pdf, doc]
        for job in jobs:
            job.result()  # raises the first failure
    print("\n".join(f"wrote {p}" for p in written))


@app.command()
def template(name: Annotated[str, typer.Argument(help=f"One of: {', '.join(docx_templates.names())}")]) -> None:
    """Regenerate templates/<name>.docx."""
    if name not in docx_templates.names():
        raise typer.BadParameter(f"no template generator {name!r}")
    out = ROOT / "templates" / f"{name}.docx"
    importlib.import_module(f"{docx_templates.__name__}.{name}").make(out)
    print(f"wrote {out}")
