import subprocess
from pathlib import Path

import pytest

from resume.chrome import headless


def test_hung_chrome_fails_instead_of_stalling(tmp_path: Path) -> None:
    """A headless Chrome that never exited once stalled the test run for over ten minutes, with no error."""
    stuck = tmp_path / "chrome"
    stuck.write_text("#!/bin/sh\nsleep 5\n")
    stuck.chmod(0o755)
    with pytest.raises(subprocess.TimeoutExpired):
        headless(stuck, "about:blank", timeout=1)
