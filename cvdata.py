# /// script
# requires-python = ">=3.12"
# dependencies = ["pyyaml~=6.0", "typer~=0.12"]
# ///
"""Load the CV for one language from the multilingual YAML.

`langs: [en, sv]` declares the languages. Anywhere in the file, a mapping holding any of those keys is a
translation point and must hold all of them, or loading fails. Its language's value is chosen: merged
into the mapping's shared keys when both are mappings (a role's facts plus its `en:`/`sv:` text), or
taking the mapping's place otherwise (`title: {en: ..., sv: ...}`). `{lang}` in `download` is filled in.
"""
from pathlib import Path

import typer
import yaml


class Untranslated(ValueError):
    """A translation point lacks a declared language, or mixes shared and translated keys badly."""


def resolve(node, lang: str, langs: set[str], path: str = ""):
    if isinstance(node, list):
        return [resolve(v, lang, langs, f"{path}[{i}]") for i, v in enumerate(node)]
    if not isinstance(node, dict):
        return node
    present = langs & node.keys()
    if not present:
        return {k: resolve(v, lang, langs, f"{path}.{k}" if path else k) for k, v in node.items()}
    if present != langs:
        raise Untranslated(f"{path or 'top level'}: has {', '.join(sorted(present))}, missing {', '.join(sorted(langs - present))}")
    shared = {k: v for k, v in node.items() if k not in langs}
    chosen = node[lang]
    if not shared:
        return resolve(chosen, lang, langs, f"{path}.{lang}")
    if not isinstance(chosen, dict):
        raise Untranslated(f"{path}.{lang}: shared keys {sorted(shared)} need a mapping per language, not a value")
    if clash := shared.keys() & chosen.keys():
        raise Untranslated(f"{path}.{lang}: {sorted(clash)} both shared and translated")
    return resolve(shared | chosen, lang, langs, path)


def langs(path: Path) -> list[str]:
    return yaml.safe_load(path.read_text(encoding="utf-8"))["langs"]


def load(path: Path, lang: str) -> dict:
    """The CV in one language, as the builders use it. Every language is checked, not just this one."""
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    declared = data.pop("langs")
    if lang not in declared:
        raise Untranslated(f"{lang!r} is not one of the declared langs {declared}")
    for other in declared:  # a gap in any language is an error, whichever one is built
        resolve(data, other, set(declared))
    cv = resolve(data, lang, set(declared)) | {"lang": lang, "_dir": path.parent}
    if "download" in cv:
        cv["download"] = cv["download"].format(lang=lang)
    return cv


app = typer.Typer(add_completion=False)


@app.command("langs")
def print_langs(path: Path) -> None:
    """The declared languages, one per line."""
    print("\n".join(langs(path)))


@app.command()
def get(path: Path, lang: str, key: str) -> None:
    """One top-level value for a language, e.g. the download name."""
    print(load(path, lang)[key])


if __name__ == "__main__":
    app()
