"""Headless Chrome: prints the web page to PDF, so the PDF is the page itself."""

import shutil
import subprocess
from pathlib import Path

CANDIDATES = ("google-chrome", "chromium", "chromium-browser", "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")


def find() -> Path | None:
    return next((Path(found) for name in CANDIDATES if (found := shutil.which(name))), None)


def headless(chrome: Path, *args: str, budget: int = 4000) -> str:
    """budget: virtual milliseconds the page's scripts get (pagination, count-ups) before Chrome acts."""
    flags = ["--headless=new", "--disable-gpu", "--window-size=1280,1000", f"--virtual-time-budget={budget}"]
    return subprocess.run([str(chrome), *flags, *args], capture_output=True, text=True, check=True).stdout


def pdf(chrome: Path, url: str, out: Path, budget: int = 6000) -> None:
    headless(chrome, "--no-pdf-header-footer", f"--print-to-pdf={out}", url, budget=budget)
