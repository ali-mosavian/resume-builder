import shutil
from pathlib import Path

import pytest

from resume import chrome as browser


@pytest.fixture
def chrome() -> Path:
    found = browser.find()
    if found is None or not shutil.which("pdftotext"):
        pytest.skip("needs Chrome and pdftotext")
    return found
