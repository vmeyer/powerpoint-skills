import argparse

import pytest

from pptx_skills.cli import main, run_handler
from pptx_skills.errors import PptxSkillsError
from pptx_skills.geometry import Box
from pptx_skills.opc import file_hash, lock_file, open_presentation, save_presentation
from pptx_skills.units import cm_to_emu, emu_to_cm, guide_to_cm


def test_units():
    assert cm_to_emu(1.5) == 540000
    assert emu_to_cm(12192000) == 33.87
    assert guide_to_cm(339) == pytest.approx(1.5, abs=0.01)


def test_box_contains_and_intersects():
    a, b = Box(0, 0, 10, 10), Box(2, 2, 3, 3)
    assert a.contains(b) and not b.contains(a)
    assert Box(0, 0, 2, 2).intersects(Box(1, 1, 2, 2)) == pytest.approx(1.0)


def test_open_potx(potx_file):
    prs = open_presentation(potx_file)
    assert len(prs.slides) == 0 and "Title Only" in [l.name for l in prs.slide_layouts]


def test_open_rejects_other_files(tmp_path):
    (tmp_path / "x.docx").write_bytes(b"")
    with pytest.raises(PptxSkillsError, match="not a .pptx or .potx"):
        open_presentation(tmp_path / "x.docx")


def test_save_returns_hash(tmp_path, default_pptx):
    digest = save_presentation(open_presentation(default_pptx), tmp_path / "o.pptx")
    assert digest == file_hash(tmp_path / "o.pptx") and len(digest) == 64


def test_lock_file(default_pptx):
    assert lock_file(default_pptx) is None
    (default_pptx.parent / f"~${default_pptx.name}").write_bytes(b"x")
    assert lock_file(default_pptx).name == f"~${default_pptx.name}"


def test_cli_exit_codes(capsys):
    assert main(["no-such-command"]) == 2
    assert main(["--help"]) == 0

    def boom(args):
        raise PptxSkillsError("bad input")

    assert run_handler(boom, argparse.Namespace()) == 2
    assert "error: bad input" in capsys.readouterr().err


def test_open_rejects_corrupt_pptx(tmp_path):
    (tmp_path / "x.pptx").write_bytes(b"")
    with pytest.raises(PptxSkillsError, match="cannot be read"):
        open_presentation(tmp_path / "x.pptx")


def test_open_rejects_corrupt_potx(tmp_path):
    (tmp_path / "x.potx").write_bytes(b"not a zip")
    with pytest.raises(PptxSkillsError, match="cannot be read"):
        open_presentation(tmp_path / "x.potx")


def test_open_rejects_missing_file(tmp_path):
    with pytest.raises(PptxSkillsError, match="does not exist"):
        open_presentation(tmp_path / "missing.pptx")


def test_open_rejects_zip_without_package_parts(tmp_path):
    import zipfile

    with zipfile.ZipFile(tmp_path / "junk.pptx", "w") as z:
        z.writestr("a.txt", "hello")
    with pytest.raises(PptxSkillsError, match="cannot be read"):
        open_presentation(tmp_path / "junk.pptx")
