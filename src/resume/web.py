"""Build a self-contained CV web page from the YAML.

The look is a theme: themes/<name>.css, plus optional <name>.js and <name>.*.js, inlined.
With more than one language, themes/_lang.js adds a language switch. Each theme prints to A4.
"""

import json
import base64
from html import escape
from pathlib import Path

from resume import ROOT
from resume.data import CV
from resume.data import Job
from resume.data import BOLD
from resume.data import span
from resume.data import when
from resume.data import Skill
from resume.data import Labels
from resume.photo import portrait

THEMES = ROOT / "themes"


def md(text: str) -> str:
    return BOLD.sub(r"<strong>\1</strong>", escape(text))


def aka(skill: Skill) -> str:
    """Other spellings the job text uses for a skill (e.g. GCP), for theme scripts that match skills to roles."""
    return f' data-aka="{escape("|".join(skill["aka"]))}"' if skill.get("aka") else ""


def kv(key: str, value: str, attrs: str = "") -> str:
    return f'<div class="kv"{attrs}><span>{escape(key)}</span><span>{escape(value)}</span></div>'


def sidebar(cv: CV, lb: Labels) -> str:
    links = "".join(f'<dd><a href="{escape(link["url"])}">{escape(link["label"])}</a></dd>' for link in cv["links"])
    skills = "".join(
        kv(s["name"], lb["levels"][s["level"]], f' style="--level:{s["level"]}"{aka(s)}') for s in cv["skills"]
    )
    langs = "".join(kv(lang["name"], lang["level"]) for lang in cv["languages"])
    return f"""<aside>
<h2>{lb["details"]}</h2><dl><dd>{escape(cv["location"])}</dd>
<dt>{lb["born"]}</dt><dd>{cv["born"]}</dd><dt>{lb["nationality"]}</dt><dd>{escape(cv["nationality"])}</dd></dl>
<h2>{lb["links"]}</h2><dl>{links}</dl>
<h2>{lb["skills"]}</h2>{skills}
<h2>{lb["languages"]}</h2>{langs}
<h2>{lb["hobbies"]}</h2><p>{escape(cv["hobbies"])}</p>
</aside>"""


def job(j: Job, lb: Labels, first_brief: bool) -> str:
    bullets = "".join(f"<li>{md(b)}</li>" for b in j["bullets"])
    tools = "".join(f"<li>{escape(t)}</li>" for t in j["tools"])
    # brief: a theme may collapse the role to one line; the first one carries the "earlier roles" caption.
    brief = " brief" if j.get("brief") else ""
    caption = f' data-earlier="{lb["earlier"]}"' if first_brief else ""
    return f"""<div class="item{brief}"{caption}><h3><span class="title">{escape(j["title"])}</span> <span class="at">{lb["at"]}</span> \
<span class="company">{escape(j["company"])}</span><span class="city">, {escape(j["city"])}</span></h3>
<div class="when">{span(j, lb)}</div><ul>{bullets}</ul><ul class="tools" aria-label="{lb["tools"]}">{tools}</ul></div>"""


def body(cv: CV, lb: Labels) -> str:
    summary = "".join(f"<p>{md(p)}</p>" for p in cv["summary"])
    first_brief = next((i for i, j in enumerate(cv["experience"]) if j.get("brief")), None)
    jobs = "".join(job(j, lb, i == first_brief) for i, j in enumerate(cv["experience"]))
    edu = "".join(
        f"""<div class="item"><h3>{escape(e["degree"])}, {escape(e["school"])}, {escape(e["city"])}</h3>
<div class="when">{span(e, lb)}</div><p>{escape(e.get("notes", ""))}</p></div>"""
        for e in cv["education"]
    )
    words = "".join(
        f"""<div class="item" tabindex="0"><h3>{escape(t["name"])}</h3><div class="when">{escape(t["role"])} · {when(t["date"], lb)}</div>
<div class="quote">{"".join(f"<blockquote>{md(q)}</blockquote>" for q in t["quote"])}</div></div>"""
        for t in cv["testimonials"]
    )
    return f"""<main>
<section class="summary"><h2>{lb["summary"]}</h2>{summary}</section>
<section class="education"><h2>{lb["education"]}</h2>{edu}</section>
<section class="experience"><h2>{lb["experience"]}</h2>{jobs}</section>
<section class="testimonials"><h2>{lb["testimonials"]}</h2>{words}</section>
</main>"""


def photo(cv: CV) -> str:
    """The YAML's photo, embedded so the page stays a single file. Themes that don't want it hide .photo."""
    if "photo" not in cv:
        return ""
    data = base64.b64encode(portrait(cv["dir"] / cv["photo"])).decode()
    return f'<img class="photo" alt="{escape(cv["name"])}" src="data:image/jpeg;base64,{data}">'


def header(cv: CV, lb: Labels) -> str:
    stats = "".join(
        f"<div><b>{escape(h['value'])}</b><span>{escape(h['label'])}</span></div>" for h in cv.get("highlights", [])
    )
    links = "".join(f'<a href="{escape(link["url"])}">{escape(link["label"])}</a>' for link in cv["links"])
    # download: basename of a published .pdf/.docx next to the page; without it, the browser's print dialog.
    if "download" in cv:
        name = escape(cv["download"])
        get = "".join(f'<a class="download" href="{name}.{ext}" download>{ext.upper()}</a>' for ext in ("pdf", "docx"))
    else:
        get = f'<button type="button" onclick="print()">{lb["print"]}</button>'
    return f"""<header>{photo(cv)}<h1>{escape(cv["name"])}</h1><div class="role">{escape(cv["title"])}</div>
<div class="where">{escape(cv["location"])}</div>
<nav class="contact">{links}{get}</nav>
<div class="highlights">{stats}</div></header>"""


def page(cv: CV, hidden: bool) -> str:
    lb = cv["labels"]
    ui = escape(json.dumps(lb.get("ui", {}), ensure_ascii=False))  # strings for theme scripts
    return f"""<div class="page" lang="{cv["lang"]}" data-ui="{ui}"{" hidden" * hidden}>
{header(cv, lb)}
<div class="grid">{sidebar(cv, lb)}{body(cv, lb)}</div>
</div>"""


def scripts(theme: str, multilingual: bool) -> list[Path]:
    """<theme>.js (often a link to the shared interactions), then any <theme>.*.js extras, in name order."""
    found = [p for p in [THEMES / f"{theme}.js", *sorted(THEMES.glob(f"{theme}.*.js"))] if p.exists()]
    return found + [THEMES / "_lang.js"] * multilingual


def render(cvs: list[CV], theme: str) -> str:
    """One page holding every language given, the first shown first."""
    js = "".join(f"<script>{p.read_text(encoding='utf-8')}</script>" for p in scripts(theme, len(cvs) > 1))
    return f"""<!doctype html>
<html lang="{cvs[0]["lang"]}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{escape(cvs[0]["name"])} - CV</title><style>{(THEMES / f"{theme}.css").read_text(encoding="utf-8")}</style></head>
<body>{"".join(page(cv, i > 0) for i, cv in enumerate(cvs))}{js}</body></html>
"""


def themes() -> list[str]:
    return sorted(p.stem for p in THEMES.glob("*.css"))
