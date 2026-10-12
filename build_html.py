# /// script
# requires-python = ">=3.12"
# dependencies = ["python-docx~=1.1", "pyyaml~=6.0", "typer~=0.12"]  # python-docx: shares date helpers with build_cv
# ///
"""Build a self-contained CV web page from the same YAML as build_cv.py.

The look is a theme: themes/<name>.css, inlined into the page. Each theme prints to A4.
"""
import re
from html import escape
from pathlib import Path

import typer
import yaml

from build_cv import span, when

THEMES = Path(__file__).with_name("themes")


def md(text: str) -> str:
    return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", escape(text))


def sidebar(cv: dict, lb: dict) -> str:
    links = "".join(f'<dd><a href="{escape(l["url"])}">{escape(l["label"])}</a></dd>' for l in cv["links"])
    skills = "".join(f'<div class="kv" style="--level:{s["level"]}"><span>{escape(s["name"])}</span><span>{lb["levels"][s["level"]]}</span></div>'
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


def job(j: dict, lb: dict) -> str:
    bullets = "".join(f"<li>{md(b)}</li>" for b in j["bullets"])
    tools = "".join(f"<li>{escape(t)}</li>" for t in j["tools"])
    return f"""<div class="item"><h3>{escape(j["title"])} {lb["at"]} {escape(j["company"])}, {escape(j["city"])}</h3>
<div class="when">{span(j, lb)}</div><ul>{bullets}</ul><ul class="tools" aria-label="{lb["tools"]}">{tools}</ul></div>"""


def body(cv: dict, lb: dict) -> str:
    summary = "".join(f"<p>{md(p)}</p>" for p in cv["summary"])
    jobs = "".join(job(j, lb) for j in cv["experience"])
    edu = "".join(f"""<div class="item"><h3>{escape(e["degree"])}, {escape(e["school"])}, {escape(e["city"])}</h3>
<div class="when">{span(e, lb)}</div><p>{escape(e.get("notes", ""))}</p></div>""" for e in cv["education"])
    words = "".join(f"""<div class="item"><h3>{escape(t["name"])}</h3><div class="when">{escape(t["role"])} · {when(t["date"], lb)}</div>
{"".join(f"<blockquote>{md(q)}</blockquote>" for q in t["quote"])}</div>""" for t in cv["testimonials"])
    return f"""<main>
<section class="summary"><h2>{lb["summary"]}</h2>{summary}</section>
<section class="experience"><h2>{lb["experience"]}</h2>{jobs}</section>
<section class="education"><h2>{lb["education"]}</h2>{edu}</section>
<section class="testimonials"><h2>{lb["testimonials"]}</h2>{words}</section>
</main>"""


def build(cv_path: Path, out: Path, theme: str = "pro") -> None:
    cv = yaml.safe_load(cv_path.read_text(encoding="utf-8"))
    lb = cv["labels"]
    out.write_text(f"""<!doctype html>
<html lang="{cv["lang"]}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{escape(cv["name"])} – CV</title><style>{(THEMES / f"{theme}.css").read_text(encoding="utf-8")}</style></head>
<body><div class="page">
<header><h1>{escape(cv["name"])}</h1><div class="role">{escape(cv["title"])}</div>
<div class="where">{escape(cv["location"])}</div></header>
<div class="grid">{sidebar(cv, lb)}{body(cv, lb)}</div>
</div></body></html>
""", encoding="utf-8")
    print(f"wrote {out}")


if __name__ == "__main__":
    typer.run(build)
