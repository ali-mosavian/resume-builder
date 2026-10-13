from pathlib import Path

import pytest
import yaml
from support import CV

from resume import data


def test_missing_language_is_an_error(tmp_path: Path) -> None:
    """One YAML holds every language: a translation point missing a declared language is an error, not a gap in the CV."""
    cv = yaml.safe_load(CV.read_text(encoding="utf-8"))
    del cv["experience"][3]["sv"]
    broken = tmp_path / "cv.yaml"
    broken.write_text(yaml.safe_dump(cv, allow_unicode=True), encoding="utf-8")
    with pytest.raises(data.Untranslated, match=r"experience\[3\].*missing sv"):
        data.load(broken, "en")


@pytest.mark.parametrize("lang", data.langs(CV))
def test_bullets_end_with_a_full_stop(lang: str) -> None:
    """A bold edit once quoted only the end of a Swedish bullet, printing a stray ' after Nordea's text."""
    for j in data.load(CV, lang)["experience"]:
        for b in j["bullets"]:
            assert b.endswith("."), f"{j['company']}: bullet ends {b[-12:]!r}"


@pytest.mark.parametrize("lang", data.langs(CV))
def test_only_the_latest_role_is_ongoing(lang: str) -> None:
    """Byggfakta once read "JUN 2024 - NU" long after it ended."""
    assert not [j["company"] for j in data.load(CV, lang)["experience"][1:] if j["end"] is None]
