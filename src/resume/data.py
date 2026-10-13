"""Load the CV for one language from the multilingual YAML.

`langs: [en, sv]` declares the languages. Anywhere in the file, a mapping holding any of those keys is a
translation point and must hold all of them, or loading fails. Its language's value is chosen: merged
into the mapping's shared keys when both are mappings (a role's facts plus its `en:`/`sv:` text), or
taking the mapping's place otherwise (`title: {en: ..., sv: ...}`). `{lang}` in `download` is filled in.
"""

import re
from pathlib import Path
from typing import NotRequired, TypedDict, cast

import yaml


class Untranslated(ValueError):
    """A translation point lacks a declared language, or mixes shared and translated keys badly."""


class Labels(TypedDict):
    details: str
    born: str
    nationality: str
    links: str
    skills: str
    languages: str
    hobbies: str
    summary: str
    experience: str
    education: str
    testimonials: str
    achievements: str
    tools: str
    at: str
    earlier: str
    print: str
    ui: dict[str, str]
    present: str
    months: list[str]
    levels: dict[int, str]


class Link(TypedDict):
    label: str
    url: str


class Skill(TypedDict):
    name: str
    level: int
    aka: NotRequired[list[str]]


class Language(TypedDict):
    name: str
    level: str


class Highlight(TypedDict):
    value: str
    label: str


class Dated(TypedDict):
    start: str
    end: str | None


class Job(Dated):
    title: str
    company: str
    city: str
    bullets: list[str]
    tools: list[str]
    brief: NotRequired[bool]


class Education(Dated):
    degree: str
    school: str
    city: str
    notes: NotRequired[str]


class Testimonial(TypedDict):
    name: str
    role: str
    date: str
    quote: list[str]


class CV(TypedDict):
    lang: str
    dir: Path
    name: str
    title: str
    location: str
    born: int
    nationality: str
    hobbies: str
    photo: NotRequired[str]
    download: NotRequired[str]
    highlights: NotRequired[list[Highlight]]
    links: list[Link]
    skills: list[Skill]
    languages: list[Language]
    summary: list[str]
    experience: list[Job]
    education: list[Education]
    testimonials: list[Testimonial]
    labels: Labels


type Node = dict[str, Node] | list[Node] | str | int | float | bool | None


def resolve(node: Node, lang: str, langs: set[str], path: str = "") -> Node:
    if isinstance(node, list):
        return [resolve(v, lang, langs, f"{path}[{i}]") for i, v in enumerate(node)]
    if not isinstance(node, dict):
        return node
    present = langs & node.keys()
    if not present:
        return {k: resolve(v, lang, langs, f"{path}.{k}" if path else k) for k, v in node.items()}
    if present != langs:
        missing = ", ".join(sorted(langs - present))
        raise Untranslated(f"{path or 'top level'}: has {', '.join(sorted(present))}, missing {missing}")
    shared = {k: v for k, v in node.items() if k not in langs}
    chosen = node[lang]
    if not shared:
        return resolve(chosen, lang, langs, f"{path}.{lang}")
    if not isinstance(chosen, dict):
        raise Untranslated(f"{path}.{lang}: shared keys {sorted(shared)} need a mapping per language, not a value")
    if clash := shared.keys() & chosen.keys():
        raise Untranslated(f"{path}.{lang}: {sorted(clash)} both shared and translated")
    return resolve(shared | chosen, lang, langs, path)


def read(path: Path) -> dict[str, Node]:
    return cast(dict[str, Node], yaml.safe_load(path.read_text(encoding="utf-8")))


def langs(path: Path) -> list[str]:
    return cast(list[str], read(path)["langs"])


def load(path: Path, lang: str) -> CV:
    """The CV in one language, as the builders use it. Every language is checked, not just this one."""
    data = read(path)
    declared = set(cast(list[str], data.pop("langs")))
    if lang not in declared:
        raise Untranslated(f"{lang!r} is not one of the declared langs {sorted(declared)}")
    for other in declared:  # a gap in any language is an error, whichever one is built
        resolve(data, other, declared)
    # ponytail: trusts the YAML's shape past translation; a typo'd key surfaces as a KeyError in a builder.
    cv = cast(CV, resolve(data, lang, declared))
    cv["lang"], cv["dir"] = lang, path.parent
    if "download" in cv:
        cv["download"] = cv["download"].format(lang=lang)
    return cv


BOLD = re.compile(r"\*\*(.+?)\*\*")


def when(ym: str | None, labels: Labels) -> str:
    if ym is None:
        return labels["present"]
    year, month = ym.split("-")
    return f"{labels['months'][int(month) - 1]} {year}"


def span(item: Dated, labels: Labels) -> str:
    return f"{when(item['start'], labels)} - {when(item['end'], labels)}".upper()
