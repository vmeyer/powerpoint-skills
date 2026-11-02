"""Repo hygiene: no Office files or images tracked, no company names, Codex manifest in sync."""

import os
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
IMAGES = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".heic", ".tif", ".tiff", ".pdf")
BINARY = (".pptx", ".potx", ".pptm", ".ppt", ".odp", *IMAGES)


def tracked_files() -> list[Path]:
    out = subprocess.run(["git", "ls-files"], cwd=ROOT, capture_output=True, text=True, check=True).stdout
    return [ROOT / f for f in out.splitlines() if (ROOT / f).is_file()]


def test_no_office_files_tracked():
    assert not [f.name for f in tracked_files() if f.suffix.lower() in BINARY]


@pytest.mark.parametrize("name", [f"shot{ext}" for ext in IMAGES] + ["master.potx", "deck.pptx", "decks/a/a.pptx", "deck-render/slide-01.png"])
def test_screenshots_and_office_files_are_ignored(name):
    """Renders and screenshots next to a corporate deck in the repo root can never be staged."""
    assert subprocess.run(["git", "check-ignore", "-q", "--no-index", name], cwd=ROOT).returncode == 0, name


@pytest.mark.skipif(not os.environ.get("PPTX_SKILLS_DENY_WORDS"), reason="local only")
def test_no_company_names():
    words = [w.strip().lower() for w in os.environ["PPTX_SKILLS_DENY_WORDS"].split(",") if w.strip()]
    for f in tracked_files():
        text = f.read_text(errors="ignore").lower()
        assert not any(w in text for w in words), f


def test_codex_manifest_in_sync():
    assert subprocess.run(["bash", "scripts/sync-mirrors.sh", "--check"], cwd=ROOT).returncode == 0
