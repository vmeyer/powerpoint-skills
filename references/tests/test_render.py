import json
import shutil
import subprocess

import pytest
from PIL import Image

from conftest import linked_deck
from pptx_skills import render as render_mod
from pptx_skills.brief import load_brief
from pptx_skills.cli import main
from pptx_skills.deckdir import init_deck, load_deck
from pptx_skills.errors import PptxSkillsError
from pptx_skills.opc import open_presentation
from pptx_skills.render import contact_sheet, doctor, render_examples, update_sheets, find_pdftoppm, find_soffice, render, render_layouts

REAL = find_soffice() is not None and find_pdftoppm() is not None
needs_tools = pytest.mark.skipif(not REAL, reason="LibreOffice/poppler not installed")


@pytest.fixture
def fake_tools(monkeypatch):
    """Record subprocess calls; fake soffice writes <stem>.pdf, fake pdftoppm writes <prefix>-N.png."""
    class Calls(list):
        kwargs: list = []

    calls = Calls()
    calls.kwargs = []

    def fake_run(cmd, **kwargs):
        calls.append(cmd)
        calls.kwargs.append(kwargs)
        if cmd[0].endswith("soffice"):
            outdir = cmd[cmd.index("--outdir") + 1]
            shutil.copyfile(cmd[-1], f"{outdir}/{cmd[-1].rsplit('/', 1)[-1].rsplit('.', 1)[0]}.pdf")
        else:
            Image.new("RGB", (160, 90), "red").save(f"{cmd[-1]}-{cmd[cmd.index('-f') + 1]}.png")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(render_mod.subprocess, "run", fake_run)
    monkeypatch.setattr(render_mod, "find_soffice", lambda: render_mod.Path("/x/soffice"))
    monkeypatch.setattr(render_mod, "find_pdftoppm", lambda: render_mod.Path("/x/pdftoppm"))
    return calls


def test_render_invokes_tools_with_isolated_profile(fake_tools, deck_with_two_slides, tmp_path):
    out = render(deck_with_two_slides, tmp_path / "out", slides=[2])
    assert any(a.startswith("-env:UserInstallation=file://") for a in fake_tools[0])
    assert all(k["timeout"] == 180 for k in fake_tools.kwargs)
    i = fake_tools[1].index("-f")
    assert fake_tools[1][i:i + 4] == ["-f", "2", "-l", "2"]
    assert [p.name for p in out] == ["slide-02.png"] and out[0].is_file()


def test_render_rejects_missing_slide(fake_tools, deck_with_two_slides, tmp_path):
    with pytest.raises(PptxSkillsError, match="slide 3 does not exist"):
        render(deck_with_two_slides, tmp_path, slides=[3])


def test_render_failure_includes_stderr_tail(monkeypatch, fake_tools, deck_with_two_slides, tmp_path):
    monkeypatch.setattr(
        render_mod.subprocess, "run", lambda cmd, **kw: subprocess.CompletedProcess(cmd, 1, "", "boom happened")
    )
    with pytest.raises(PptxSkillsError, match="boom happened"):
        render(deck_with_two_slides, tmp_path)


def test_render_timeout(monkeypatch, fake_tools, deck_with_two_slides, tmp_path):
    def slow(cmd, **kw):
        raise subprocess.TimeoutExpired(cmd, 180)

    monkeypatch.setattr(render_mod.subprocess, "run", slow)
    with pytest.raises(PptxSkillsError, match="timed out"):
        render(deck_with_two_slides, tmp_path)


def test_missing_soffice_names_install_command(monkeypatch, deck_with_two_slides):
    monkeypatch.setattr(render_mod, "find_soffice", lambda: None)
    with pytest.raises(PptxSkillsError, match="brew install --cask libreoffice") as exc:
        render(deck_with_two_slides)
    assert "brew install poppler" in str(exc.value) and "apt install libreoffice poppler-utils" in str(exc.value)


def test_find_soffice_falls_back_to_app_folders(monkeypatch, tmp_path):
    app = tmp_path / "soffice"
    app.touch()
    monkeypatch.setattr(render_mod.shutil, "which", lambda name: None)
    monkeypatch.setattr(render_mod, "_SOFFICE_PATHS", (tmp_path / "none", app))
    assert find_soffice() == app


def test_default_out_dir(tmp_path, default_pptx):
    deck_dir = tmp_path / "decks/d"
    init_deck(deck_dir, default_pptx)
    assert render_mod.default_out_dir(deck_dir / "d.pptx") == deck_dir / "render"
    assert render_mod.default_out_dir(default_pptx) == default_pptx.parent / "default-render"


def test_contact_sheet_size(tmp_path):
    pngs = []
    for i in range(5):
        pngs.append(tmp_path / f"{i}.png")
        Image.new("RGB", (1600, 900), "blue").save(pngs[-1])
    out = contact_sheet(pngs, list("abcde"), tmp_path / "sheet.png")
    with Image.open(out) as sheet:
        gap = render_mod._GAP
        assert sheet.width == 4 * 480 + 5 * gap
        assert sheet.height == 2 * (270 + render_mod._LABEL_H) + 3 * gap


def test_doctor_reports_fonts_and_exit_code(monkeypatch, tmp_path, default_pptx, capsys):
    deck_dir = tmp_path / "decks/d"
    init_deck(deck_dir, default_pptx)
    monkeypatch.setattr(render_mod, "find_soffice", lambda: render_mod.Path("/x/soffice"))
    monkeypatch.setattr(render_mod, "find_pdftoppm", lambda: render_mod.Path("/x/pdftoppm"))
    monkeypatch.setattr(render_mod, "find_font", lambda family: None)
    report = doctor(deck_dir)
    assert report["soffice"] == "/x/soffice" and report["fonts"] and not any(report["fonts"].values())
    assert main(["doctor", str(deck_dir)]) == 1
    assert "✗ font" in capsys.readouterr().out
    monkeypatch.setattr(render_mod, "find_font", lambda family: render_mod.Path("/f.ttf"))
    assert main(["doctor", str(deck_dir), "--json"]) == 0
    assert json.loads(capsys.readouterr().out)["ok"] is True


def test_init_without_libreoffice_still_succeeds(monkeypatch, tmp_path, default_pptx, capsys):
    monkeypatch.setattr(render_mod, "find_soffice", lambda: None)
    assert main(["init", str(tmp_path / "d"), "--from", str(default_pptx)]) == 0
    assert "contact sheets skipped: LibreOffice not found — run `pptx doctor`" in capsys.readouterr().out
    assert not (tmp_path / "d/style/layouts.png").exists()


def test_style_stores_hash_and_skips_unchanged_sheets(monkeypatch, tmp_path, default_pptx):
    rendered = []

    def fake_layouts(deck, out):
        rendered.append(out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(b"png")

    monkeypatch.setattr(render_mod, "find_soffice", lambda: render_mod.Path("/x/soffice"))
    monkeypatch.setattr(render_mod, "find_pdftoppm", lambda: render_mod.Path("/x/pdftoppm"))
    monkeypatch.setattr(render_mod, "render_layouts", fake_layouts)
    d = tmp_path / "d"
    init_deck(d, default_pptx)
    assert len(rendered) == 1 and load_brief(d)["sources_hash"]
    assert main(["style", str(d)]) == 0
    assert len(rendered) == 1  # unchanged sources: skipped
    assert main(["style", str(d), "--refresh"]) == 0
    assert len(rendered) == 2
    deck = load_deck(d)
    (d / deck["working_copy"]).write_bytes((d / deck["working_copy"]).read_bytes() + b"\0")
    assert main(["style", str(d)]) == 0
    assert len(rendered) == 3


@needs_tools
def test_render_real(deck_with_two_slides, tmp_path):
    out = render(deck_with_two_slides, tmp_path)
    assert [p.name for p in out] == ["slide-01.png", "slide-02.png"]
    with Image.open(out[0]) as image:
        assert image.width > 500


@needs_tools
def test_cli_render_sheet_and_layout_sheet_real(deck_with_two_slides, tmp_path, capsys):
    out_dir = tmp_path / "r"
    assert main(["render", str(deck_with_two_slides), "--slides", "2", "--dpi", "50", "--sheet", "--out", str(out_dir)]) == 0
    assert sorted(p.name for p in out_dir.iterdir()) == ["sheet.png", "slide-02.png"]
    sheet = render_layouts(deck_with_two_slides, tmp_path / "layouts.png")
    with Image.open(sheet) as image:
        assert image.width > 1000


def test_unwritable_out_dir_is_a_pptx_error(fake_tools, deck_with_two_slides, tmp_path):
    blocker = tmp_path / "file"
    blocker.write_text("x")
    with pytest.raises(PptxSkillsError, match="cannot create output folder"):
        render(deck_with_two_slides, blocker / "sub")


def test_corrupt_png_is_a_pptx_error(tmp_path):
    bad = tmp_path / "bad.png"
    bad.write_bytes(b"not a png")
    with pytest.raises(PptxSkillsError, match="cannot read image"):
        contact_sheet([bad], ["x"], tmp_path / "s.png")


def test_unwritable_sheet_is_a_pptx_error(tmp_path):
    good = tmp_path / "g.png"
    Image.new("RGB", (10, 10)).save(good)
    blocker = tmp_path / "file"
    blocker.write_text("x")
    with pytest.raises(PptxSkillsError, match="cannot write contact sheet"):
        contact_sheet([good], ["x"], blocker / "s.png")


def test_init_survives_sheet_write_failure(monkeypatch, tmp_path, default_pptx, capsys):
    def boom(deck, out):
        raise OSError("disk full")

    monkeypatch.setattr(render_mod, "find_soffice", lambda: render_mod.Path("/x/soffice"))
    monkeypatch.setattr(render_mod, "find_pdftoppm", lambda: render_mod.Path("/x/pdftoppm"))
    monkeypatch.setattr(render_mod, "render_layouts", boom)
    assert main(["init", str(tmp_path / "d"), "--from", str(default_pptx)]) == 0
    assert "contact sheets skipped: disk full" in capsys.readouterr().out


def test_slides_option_must_be_numbers(deck_with_two_slides, capsys):
    assert main(["render", str(deck_with_two_slides), "--slides", "abc"]) == 2
    assert "--slides expects numbers" in capsys.readouterr().err


def test_temp_dirs_removed_after_failed_conversion(monkeypatch, fake_tools, deck_with_two_slides, tmp_path):
    tmp_root = tmp_path / "tmproot"
    tmp_root.mkdir()
    monkeypatch.setattr(render_mod.tempfile, "tempdir", str(tmp_root))
    monkeypatch.setattr(
        render_mod.subprocess, "run", lambda cmd, **kw: subprocess.CompletedProcess(cmd, 1, "", "nope")
    )
    with pytest.raises(PptxSkillsError):
        render(deck_with_two_slides, tmp_path / "o")
    assert list(tmp_root.iterdir()) == []


def test_render_examples_and_sheet_hash_paths(monkeypatch, fake_tools, deck_with_two_slides, tmp_path, default_pptx):
    labels = []
    real_sheet = render_mod.contact_sheet

    def spy(pngs, names, out, **kw):
        labels.append(names)
        return real_sheet(pngs, names, out, **kw)

    monkeypatch.setattr(render_mod, "contact_sheet", spy)
    out = render_examples(deck_with_two_slides, tmp_path / "ex.png")
    assert out.is_file() and labels[-1] == ["Slide 1", "Slide 2"]

    d = tmp_path / "d"
    init_deck(d, default_pptx, examples=deck_with_two_slides)
    assert (d / "style/examples.png").is_file() and (d / "style/layouts.png").is_file()
    deck = load_deck(d)
    h = load_brief(d)["sources_hash"]
    assert h
    n = len(fake_tools)
    assert update_sheets(d, deck, h) == (h, []) and len(fake_tools) == n  # unchanged: nothing rendered
    assert update_sheets(d, deck, h, force=True)[0] == h and len(fake_tools) > n
    (d / "style/examples.png").unlink()
    n = len(fake_tools)
    update_sheets(d, deck, h)
    assert (d / "style/examples.png").is_file() and len(fake_tools) > n


def test_layout_sheet_of_a_deck_whose_zoom_follows_its_target(fake_tools, make_deck, tmp_path):
    """The deck's own slides are all dropped, so a zoom must not block dropping its target."""
    deck = linked_deck(make_deck, source=3, target=1, zoom=True)
    sheet = render_layouts(deck, tmp_path / "layouts.png")
    assert sheet.is_file()
    pages = [cmd for cmd in fake_tools if not cmd[0].endswith("soffice")]
    assert len(pages) == len(open_presentation(deck).slide_layouts)  # one page per layout, no deck slide left
