# /// script
# requires-python = ">=3.12"
# dependencies = ["python-docx~=1.1", "pyyaml~=6.0", "typer~=0.12"]  # python-docx: shares date helpers with build_cv
# ///
"""Build a self-contained CV web page from the same YAML as build_cv.py.

The look is a theme: themes/<name>.css, plus optional <name>.js and <name>.*.js, inlined. Each theme prints to A4.
"""
import json
import re
from html import escape
from pathlib import Path
from typing import Annotated

import typer
import yaml

from build_cv import span, when

THEMES = Path(__file__).with_name("themes")


def md(text: str) -> str:
    return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", escape(text))


def aka(skill: dict) -> str:
    """Other spellings the job text uses for a skill (e.g. GCP), for theme scripts that match skills to roles."""
    return f' data-aka="{escape("|".join(skill["aka"]))}"' if skill.get("aka") else ""


def sidebar(cv: dict, lb: dict) -> str:
    links = "".join(f'<dd><a href="{escape(l["url"])}">{escape(l["label"])}</a></dd>' for l in cv["links"])
    skills = "".join(f'<div class="kv" style="--level:{s["level"]}"{aka(s)}><span>{escape(s["name"])}</span><span>{lb["levels"][s["level"]]}</span></div>'
                     for s in cv["skills"])
    langs = "".join(f'<div class="kv"><span>{escape(l["name"])}</span><span>{escape(l["level"])}</span></div>'
                    for l in cv["languages"])
    return f"""<aside>
<h2>{lb["details"]}</h2><dl><dd>{escape(cv["location"])}</dd>
<dt>{lb["born"]}</dt><dd>{cv["born"]}</dd><dt>{lb["nationality"]}</dt><dd>{escape(cv["nationality"])}</dd></dl>
<h2>{lb["links"]}</h2><dl>{links}</dl>
<h2>{lb["skills"]}</h2>{skills}
<h2>{lb["languages"]}</h2>{langs}
<h2>{lb["hobbies"]}</h2><p>{escape(cv["hobbies"])}</p>
</aside>"""


def job(j: dict, lb: dict, first_brief: bool) -> str:
    bullets = "".join(f"<li>{md(b)}</li>" for b in j["bullets"])
    tools = "".join(f"<li>{escape(t)}</li>" for t in j["tools"])
    # brief: a theme may collapse the role to one line; the first one carries the "earlier roles" caption.
    brief = ' brief' if j.get("brief") else ''
    caption = f' data-earlier="{lb["earlier"]}"' if first_brief else ''
    return f"""<div class="item{brief}"{caption}><h3><span class="title">{escape(j["title"])}</span> <span class="at">{lb["at"]}</span> \
<span class="company">{escape(j["company"])}</span><span class="city">, {escape(j["city"])}</span></h3>
<div class="when">{span(j, lb)}</div><ul>{bullets}</ul><ul class="tools" aria-label="{lb["tools"]}">{tools}</ul></div>"""


def body(cv: dict, lb: dict) -> str:
    summary = "".join(f"<p>{md(p)}</p>" for p in cv["summary"])
    first_brief = next((i for i, j in enumerate(cv["experience"]) if j.get("brief")), None)
    jobs = "".join(job(j, lb, i == first_brief) for i, j in enumerate(cv["experience"]))
    edu = "".join(f"""<div class="item"><h3>{escape(e["degree"])}, {escape(e["school"])}, {escape(e["city"])}</h3>
<div class="when">{span(e, lb)}</div><p>{escape(e.get("notes", ""))}</p></div>""" for e in cv["education"])
    words = "".join(f"""<div class="item" tabindex="0"><h3>{escape(t["name"])}</h3><div class="when">{escape(t["role"])} · {when(t["date"], lb)}</div>
<div class="quote">{"".join(f"<blockquote>{md(q)}</blockquote>" for q in t["quote"])}</div></div>""" for t in cv["testimonials"])
    return f"""<main>
<section class="summary"><h2>{lb["summary"]}</h2>{summary}</section>
<section class="experience"><h2>{lb["experience"]}</h2>{jobs}</section>
<section class="education"><h2>{lb["education"]}</h2>{edu}</section>
<section class="testimonials"><h2>{lb["testimonials"]}</h2>{words}</section>
</main>"""


def header(cv: dict, lb: dict) -> str:
    stats = "".join(f'<div><b>{escape(h["value"])}</b><span>{escape(h["label"])}</span></div>' for h in cv.get("highlights", []))
    links = "".join(f'<a href="{escape(l["url"])}">{escape(l["label"])}</a>' for l in cv["links"])
    return f"""<header><h1>{escape(cv["name"])}</h1><div class="role">{escape(cv["title"])}</div>
<div class="where">{escape(cv["location"])}</div>
<nav class="contact">{links}<button type="button" onclick="print()">{lb["print"]}</button></nav>
<div class="highlights">{stats}</div></header>"""


def page(cv: dict, hidden: bool) -> str:
    lb = cv["labels"]
    ui = escape(json.dumps(lb.get("ui", {}), ensure_ascii=False))  # strings for theme scripts
    return f"""<div class="page" lang="{cv["lang"]}" data-ui="{ui}"{" hidden" * hidden}>
{header(cv, lb)}
<div class="grid">{sidebar(cv, lb)}{body(cv, lb)}</div>
</div>"""


def build(cv_path: Path, out: Path, theme: str = "pro",
          also: Annotated[list[Path] | None, typer.Option(help="More YAMLs (other languages); a theme script may switch.")] = None) -> None:
    cvs = [yaml.safe_load(p.read_text(encoding="utf-8")) for p in (cv_path, *(also or ()))]
    # <theme>.js (often a link to the shared interactions), then any <theme>.*.js extras, in name order.
    scripts = [p for p in [THEMES / f"{theme}.js", *sorted(THEMES.glob(f"{theme}.*.js"))] if p.exists()]
    js = "".join(f"<script>{p.read_text(encoding='utf-8')}</script>" for p in scripts)
    out.write_text(f"""<!doctype html>
<html lang="{cvs[0]["lang"]}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{escape(cvs[0]["name"])} – CV</title><style>{(THEMES / f"{theme}.css").read_text(encoding="utf-8")}</style></head>
<body>{"".join(page(cv, i > 0) for i, cv in enumerate(cvs))}{js}</body></html>
""", encoding="utf-8")
    print(f"wrote {out}")


if __name__ == "__main__":
    typer.run(build)
