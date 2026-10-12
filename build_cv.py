# /// script
# requires-python = ">=3.12"
# dependencies = ["python-docx~=1.1", "pyyaml~=6.0", "typer~=0.12"]
# ///
"""Build a CV .docx from a YAML file, styled by a .docx template.

The template's single table is the layout: column 0 is the sidebar, column 2 the main body.
Restyle the template in Word; paragraph styles are referenced by style id below.
"""
import re
from pathlib import Path

import typer
import yaml
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BOLD = re.compile(r"\*\*(.+?)\*\*")


def add_text(p, text: str) -> None:
    """Append text with **bold** markup as runs."""
    for i, part in enumerate(BOLD.split(text)):
        if part:
            # None, not False, outside markup: False would override bold paragraph styles.
            p.add_run(part).bold = True if i % 2 else None


def para(cell, style_id: str, text: str = ""):
    # By style id, not name: the resume.io template has duplicate style names.
    p = cell.add_paragraph()
    p._p.style = style_id
    add_text(p, text)
    return p


def section(cell, spacing_id: str, heading_id: str, title: str) -> None:
    para(cell, spacing_id)
    para(cell, heading_id, title)


def bullet_num_id(doc) -> str:
    """numId of the template's first bullet list, so bullets match the template."""
    numbering = doc.part.numbering_part.element
    bullets = {a.get(qn("w:abstractNumId")) for a in numbering.findall(qn("w:abstractNum"))
               if a.find(f"{qn('w:lvl')}/{qn('w:numFmt')}").get(qn("w:val")) == "bullet"}
    return next(n.get(qn("w:numId")) for n in numbering.findall(qn("w:num"))
                if n.find(qn("w:abstractNumId")).get(qn("w:val")) in bullets)


def bullet(cell, text: str, num_id: str) -> None:
    p = para(cell, "ListParagraph", text)
    num_pr = OxmlElement("w:numPr")
    for tag, val in (("w:ilvl", "0"), ("w:numId", num_id)):
        el = OxmlElement(tag)
        el.set(qn("w:val"), val)
        num_pr.append(el)
    p._p.get_or_add_pPr().append(num_pr)


def hyperlink(cell, label: str, url: str) -> None:
    p = para(cell, "SidebarText")
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), p.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True))
    run = p.add_run(label)
    run.underline = True
    link.append(run._r)
    p._p.append(link)


def when(ym: str | None, labels: dict) -> str:
    if ym is None:
        return labels["present"]
    year, month = str(ym).split("-")
    return f"{labels['months'][int(month) - 1]} {year}"


def span(item: dict, labels: dict) -> str:
    return f"{when(item['start'], labels)} — {when(item['end'], labels)}".upper()


def header(paragraphs, cv: dict) -> None:
    """Fill the template's Name / JobTitle placeholders, wherever the template puts them."""
    for p in paragraphs:
        if p._p.style in ("Name", "JobTitle"):
            p.runs[0].text = cv["name"] if p._p.style == "Name" else cv["title"].upper()


def sidebar(cell, cv: dict, lb: dict) -> None:
    section(cell, "SidebarSectionsSpacing", "Heading3", lb["details"])
    para(cell, "SidebarText", cv["location"])
    for key in ("born", "nationality"):
        para(cell, "Heading4", lb[key].upper())
        para(cell, "SidebarText", str(cv[key]))
    section(cell, "SidebarSectionsSpacing", "Heading3", lb["links"])
    for link in cv["links"]:
        hyperlink(cell, link["label"], link["url"])
    section(cell, "SidebarSectionsSpacing", "Heading3", lb["skills"])
    for s in cv["skills"]:
        para(cell, "SidebarText", f"{s['name']} · {lb['levels'][s['level']]}")
    section(cell, "SidebarSectionsSpacing", "Heading3", lb["languages"])
    for lang in cv["languages"]:
        para(cell, "SidebarText", f"{lang['name']} · {lang['level']}")
    section(cell, "SidebarSectionsSpacing", "Heading3", lb["hobbies"])
    para(cell, "SidebarText", cv["hobbies"])


def main_body(cell, cv: dict, lb: dict, num_id: str) -> None:
    cell._tc.remove(cell.paragraphs[0]._p)  # template placeholder
    para(cell, "Heading1", lb["summary"])
    for text in cv["summary"]:
        para(cell, "Normal", text)
    section(cell, "MainSectionsSpacing", "Heading1", lb["experience"])
    for job in cv["experience"]:
        para(cell, "Heading2", f"{job['title']} {lb['at']} {job['company']}, {job['city']}")
        para(cell, "Date", span(job, lb))
        para(cell, "Normal").add_run(lb["achievements"]).italic = True
        for text in job["bullets"]:
            bullet(cell, text, num_id)
        p = para(cell, "Normal")
        p.add_run(f"{lb['tools']}: ").italic = True
        p.add_run(", ".join(job["tools"]))
    section(cell, "MainSectionsSpacing", "Heading1", lb["education"])
    for ed in cv["education"]:
        para(cell, "Heading2", f"{ed['degree']}, {ed['school']}, {ed['city']}")
        para(cell, "Date", span(ed, lb))
        para(cell, "Normal", ed.get("notes", ""))
    section(cell, "MainSectionsSpacing", "Heading1", lb["testimonials"])
    for t in cv["testimonials"]:
        para(cell, "Heading2", f"{t['name']} | {t['role']}")
        para(cell, "Date", when(t["date"], lb).upper())
        quote = [*t["quote"]]
        quote[0] = f"“{quote[0]}"
        quote[-1] += "”"
        for text in quote:
            para(cell, "Normal", text)


def build(cv_path: Path, out: Path, template: Path = Path(__file__).with_name("template.docx")) -> None:
    cv = yaml.safe_load(cv_path.read_text(encoding="utf-8"))
    doc = Document(template)
    side, _, main = doc.tables[0].rows[0].cells
    header([*doc.paragraphs, *side.paragraphs], cv)
    sidebar(side, cv, cv["labels"])
    main_body(main, cv, cv["labels"], bullet_num_id(doc))
    doc.save(out)
    print(f"wrote {out}")


if __name__ == "__main__":
    typer.run(build)
