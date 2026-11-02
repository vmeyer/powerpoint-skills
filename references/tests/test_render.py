import json
import os
import plistlib
import re
import shutil
import subprocess
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image
from pptx import Presentation

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


def test_init_without_a_renderer_still_succeeds(monkeypatch, tmp_path, default_pptx, capsys):
    monkeypatch.setattr(render_mod, "find_soffice", lambda: None)
    assert main(["init", str(tmp_path / "d"), "--from", str(default_pptx), "--sheets"]) == 0
    assert "contact sheets skipped: neither PowerPoint nor LibreOffice found" in capsys.readouterr().out
    assert not (tmp_path / "d/style/layouts.png").exists()


def test_style_stores_hash_and_skips_unchanged_sheets(monkeypatch, tmp_path, default_pptx):
    rendered = []

    def fake_layouts(deck, out, *args):
        rendered.append(out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(b"png")

    monkeypatch.setattr(render_mod, "find_soffice", lambda: render_mod.Path("/x/soffice"))
    monkeypatch.setattr(render_mod, "find_pdftoppm", lambda: render_mod.Path("/x/pdftoppm"))
    monkeypatch.setattr(render_mod, "render_layouts", fake_layouts)
    d = tmp_path / "d"
    init_deck(d, default_pptx, sheets=True)
    assert len(rendered) == 1 and load_brief(d)["sources_hash"]
    assert main(["style", str(d), "--sheets"]) == 0
    assert len(rendered) == 1  # unchanged sources: skipped
    assert main(["style", str(d), "--refresh"]) == 0
    assert len(rendered) == 2
    deck = load_deck(d)
    (d / deck["working_copy"]).write_bytes((d / deck["working_copy"]).read_bytes() + b"\0")
    assert main(["style", str(d)]) == 0  # without --sheets: the brief only, and the hash stays
    assert len(rendered) == 2
    assert main(["style", str(d), "--sheets"]) == 0
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
    def boom(deck, out, *args):
        raise OSError("disk full")

    monkeypatch.setattr(render_mod, "find_soffice", lambda: render_mod.Path("/x/soffice"))
    monkeypatch.setattr(render_mod, "find_pdftoppm", lambda: render_mod.Path("/x/pdftoppm"))
    monkeypatch.setattr(render_mod, "render_layouts", boom)
    assert main(["init", str(tmp_path / "d"), "--from", str(default_pptx), "--sheets"]) == 0
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
    init_deck(d, default_pptx, examples=deck_with_two_slides, sheets=True)
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


@needs_tools
def test_neutral_deck_renders_layout_sheet(tmp_path):
    style = tmp_path / "style.json"
    style.write_text("{}")
    init_deck(tmp_path / "n", None, style_file=style, sheets=True)
    assert (tmp_path / "n/style/layouts.png").stat().st_size > 0


def hidden_deck(make_deck):
    """Three blank slides, each a full-slide rectangle in its own color (red, green, blue); slide 2 is hidden."""
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE

    from conftest import layout_named

    def build(prs):
        for rgb in ("FF0000", "00FF00", "0000FF"):
            slide = prs.slides.add_slide(layout_named(prs, "Blank"))
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(rgb)
        prs.slides[1]._element.set("show", "0")

    return make_deck(build, name="hidden.pptx")


def test_render_exports_hidden_slides(fake_tools, deck_with_two_slides, tmp_path):
    render(deck_with_two_slides, tmp_path / "out")
    i = fake_tools[0].index("--convert-to")
    assert fake_tools[0][i + 1] == 'pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}'


@needs_tools
def test_render_real_with_a_hidden_slide(make_deck, tmp_path):
    deck = hidden_deck(make_deck)
    out = render(deck, tmp_path / "all", dpi=20)
    assert [p.name for p in out] == ["slide-01.png", "slide-02.png", "slide-03.png"]
    colors = []
    for path in out:
        with Image.open(path) as image:
            colors.append(image.convert("RGB").getpixel((image.width // 2, image.height // 2)))
    assert colors == [(255, 0, 0), (0, 255, 0), (0, 0, 255)]  # page n is slide n, hidden or not
    last = render(deck, tmp_path / "last", slides=[3], dpi=20)
    with Image.open(last[0]) as image:
        assert image.convert("RGB").getpixel((image.width // 2, image.height // 2)) == (0, 0, 255)


# --- PowerPoint as the preferred renderer -------------------------------------------------------


@pytest.fixture
def fake_powerpoint(monkeypatch, tmp_path):
    """An installed PowerPoint (version 16.99) driven by a fake osascript; LibreOffice is missing unless a test adds it.

    `state.behavior` decides what an export does: "ok" writes the PDF, "fail" exits 1 as an Automation denial does,
    "consent" exits 1 as a refused automation consent does, "cancel" exits 1 as a cancelled dialog does,
    "timeout" never answers, "silent" exits 0 without a PDF. Every export records the copy PowerPoint was given and,
    like Office, leaves owner files next to it: `~$<copy>` and a truncated `~$er-<hex>.pptx`.
    """
    app = tmp_path / "Applications/Microsoft PowerPoint.app"
    (app / "Contents").mkdir(parents=True)
    (app / "Contents/Info.plist").write_bytes(plistlib.dumps({"CFBundleShortVersionString": "16.99"}))
    monkeypatch.setattr(render_mod, "_POWERPOINT_PATHS", (tmp_path / "missing/Microsoft PowerPoint.app", app))
    monkeypatch.setattr(render_mod, "_MACOS", True)
    monkeypatch.setattr(render_mod, "find_soffice", lambda: None)
    monkeypatch.setattr(render_mod, "find_pdftoppm", lambda: Path("/x/pdftoppm"))
    state = SimpleNamespace(app=app, behavior="ok", exports=[], closes=[], soffice=[], pdftoppm=[])

    def fake_run(cmd, **kwargs):
        if cmd[0] == "osascript":
            script, args = cmd[2], cmd[3:]
            if script == render_mod._CLOSE_SCRIPT:
                state.closes.append(args[0])
                return subprocess.CompletedProcess(cmd, 0, "", "")
            copy, pdf, stem = Path(args[0]), Path(args[1]), args[2]
            shown = [slide._element.get("show") for slide in Presentation(str(copy)).slides]
            state.exports.append({"copy": copy, "pdf": pdf, "stem": stem, "script": script, "shown": shown, "timeout": kwargs.get("timeout")})
            (copy.parent / f"~${copy.name}").write_bytes(b"owner")
            (copy.parent / f"~${copy.name[4:]}").write_bytes(b"owner")
            if state.behavior == "timeout":
                raise subprocess.TimeoutExpired(cmd, kwargs["timeout"])
            if state.behavior == "fail":
                return subprocess.CompletedProcess(cmd, 1, "", "execution error: Not authorized to send Apple events to Microsoft PowerPoint. (-1743)")
            if state.behavior == "consent":
                return subprocess.CompletedProcess(cmd, 1, "", "execution error: Not authorized to send Apple events to Microsoft PowerPoint. (-1744)")
            if state.behavior == "cancel":
                return subprocess.CompletedProcess(cmd, 1, "", "execution error: Microsoft PowerPoint got an error: User canceled. (-128)")
            if state.behavior == "ok":
                shutil.copyfile(copy, pdf)
            return subprocess.CompletedProcess(cmd, 0, "", "")
        if cmd[0].endswith("soffice"):
            state.soffice.append(cmd)
            outdir = cmd[cmd.index("--outdir") + 1]
            shutil.copyfile(cmd[-1], f"{outdir}/{Path(cmd[-1]).stem}.pdf")
        else:
            state.pdftoppm.append(cmd)
            Image.new("RGB", (160, 90), "red").save(f"{cmd[-1]}-{cmd[cmd.index('-f') + 1]}.png")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(render_mod.subprocess, "run", fake_run)
    return state


def with_libreoffice(monkeypatch):
    monkeypatch.setattr(render_mod, "find_soffice", lambda: Path("/x/soffice"))


def hidden_entries(folder):
    return [p for p in Path(folder).rglob(".*")]


def assert_empty_export_folder(folder):
    """PowerPoint's exchange folder `<folder>/_export` stays (its one-time grant keeps working) and is empty again."""
    export = Path(folder) / "_export"
    assert export.is_dir() and list(export.iterdir()) == []


def export_folders(state) -> set:
    return {export["copy"].parent.resolve() for export in state.exports}


def test_auto_prefers_powerpoint_then_libreoffice(fake_powerpoint, monkeypatch):
    with_libreoffice(monkeypatch)
    assert render_mod.select_renderer("auto") == "powerpoint"
    monkeypatch.setattr(render_mod, "_MACOS", False)  # PowerPoint renders on macOS only
    assert render_mod.select_renderer("auto") == "libreoffice"
    monkeypatch.setattr(render_mod, "_MACOS", True)
    monkeypatch.setattr(render_mod, "_POWERPOINT_PATHS", ())
    assert render_mod.select_renderer("auto") == "libreoffice"
    monkeypatch.setattr(render_mod, "find_soffice", lambda: None)
    assert render_mod.select_renderer("auto") is None


def test_find_powerpoint_in_the_user_applications_folder_and_read_its_version(fake_powerpoint):
    assert render_mod.find_powerpoint() == fake_powerpoint.app
    assert render_mod.powerpoint_version(fake_powerpoint.app) == "16.99"
    assert render_mod.powerpoint_version(fake_powerpoint.app.parent / "Other.app") is None


def test_renderer_option_beats_the_environment_which_beats_auto(monkeypatch):
    assert render_mod.requested_renderer() == "auto"
    monkeypatch.setenv("PPTX_SKILLS_RENDERER", "LibreOffice")
    assert render_mod.requested_renderer() == "libreoffice"
    assert render_mod.requested_renderer("powerpoint") == "powerpoint"
    monkeypatch.setenv("PPTX_SKILLS_RENDERER", "keynote")
    with pytest.raises(PptxSkillsError, match="PPTX_SKILLS_RENDERER"):
        render_mod.requested_renderer()


def test_a_forced_renderer_must_be_installed(fake_powerpoint, monkeypatch, deck_with_two_slides, tmp_path):
    with pytest.raises(PptxSkillsError, match=r"LibreOffice \(soffice\) not found.*brew install --cask libreoffice"):
        render(deck_with_two_slides, tmp_path / "o", renderer="libreoffice")
    monkeypatch.setattr(render_mod, "_POWERPOINT_PATHS", ())
    with_libreoffice(monkeypatch)
    monkeypatch.setenv("PPTX_SKILLS_RENDERER", "powerpoint")
    with pytest.raises(PptxSkillsError, match="PowerPoint not found"):
        render(deck_with_two_slides, tmp_path / "o")
    assert fake_powerpoint.soffice == []


def test_powerpoint_renders_an_unhidden_copy_in_the_export_folder_next_to_the_deck(fake_powerpoint, make_deck, tmp_path):
    deck = hidden_deck(make_deck)
    out = tmp_path / "out"
    files = render(deck, out, dpi=77)
    assert [p.name for p in files] == ["slide-01.png", "slide-02.png", "slide-03.png"]
    (export,) = fake_powerpoint.exports
    # PowerPoint writes only into the exact folders the user granted: one fixed, visible `_export` next to the deck
    assert export["copy"].parent == deck.parent / "_export"
    assert re.fullmatch(r"render-[0-9a-f]{32}", export["stem"]) and export["copy"].name == f"{export['stem']}.pptx"
    assert export["pdf"].parent == export["copy"].parent
    assert export["shown"] == [None, None, None]  # PowerPoint's PDF would leave the hidden slide 2 out
    assert open_presentation(deck).slides[1]._element.get("show") == "0"  # the deck itself is untouched
    assert export["timeout"] == 120
    assert [cmd[cmd.index("-r") + 1] for cmd in fake_powerpoint.pdftoppm] == ["77"] * 3
    assert sorted(p.name for p in out.iterdir()) == ["slide-01.png", "slide-02.png", "slide-03.png"]
    assert_empty_export_folder(deck.parent)
    assert hidden_entries(tmp_path) == []
    assert fake_powerpoint.closes == [] and fake_powerpoint.soffice == []


def test_powerpoint_copies_are_unique_and_the_scripts_touch_nothing_else(fake_powerpoint, deck_with_two_slides, tmp_path):
    render(deck_with_two_slides, tmp_path / "out")
    render(deck_with_two_slides, tmp_path / "out", slides=[2])
    first, second = fake_powerpoint.exports
    assert first["stem"] != second["stem"]
    assert first["copy"].parent == second["copy"].parent == tmp_path / "_export"  # one folder, granted once
    assert_empty_export_folder(tmp_path)
    assert first["script"] == render_mod._EXPORT_SCRIPT and str(first["copy"]) not in first["script"]  # paths are arguments
    for script in (render_mod._EXPORT_SCRIPT, render_mod._CLOSE_SCRIPT):
        assert "quit" not in script and "active presentation" not in script
    assert "is running" in render_mod._CLOSE_SCRIPT  # cleaning up never launches PowerPoint


_AUTOMATION = ("Privacy & Security → Automation",)
_GRANT = ('"Grant Access" dialog for {export}', 'click "Grant Access"', "once per project", "fall back to LibreOffice in auto mode")


@pytest.mark.parametrize("behavior, needles", [
    ("fail", ("-1743", *_AUTOMATION)),
    ("consent", ("-1744", *_AUTOMATION)),
    ("timeout", ("within 120 s", "may control Microsoft PowerPoint", *_GRANT)),
    ("silent", ("no PDF", *_GRANT)),
    ("cancel", ("-128", *_AUTOMATION, *_GRANT)),
])
def test_powerpoint_failure_without_libreoffice_explains_the_setup(fake_powerpoint, deck_with_two_slides, tmp_path, behavior, needles):
    fake_powerpoint.behavior = behavior
    out = tmp_path / "out"
    with pytest.raises(PptxSkillsError) as exc:
        render(deck_with_two_slides, out)
    message = str(exc.value)
    for needle in needles:
        assert needle.format(export=tmp_path / "_export") in message
    assert list(out.iterdir()) == []
    assert_empty_export_folder(tmp_path)  # copy, owner file and PDF are gone, the folder stays
    assert fake_powerpoint.closes == [fake_powerpoint.exports[0]["stem"]]  # only the copy is closed


@pytest.mark.parametrize("behavior", ["fail", "timeout", "silent"])
def test_powerpoint_failure_falls_back_to_libreoffice_with_a_note(fake_powerpoint, monkeypatch, deck_with_two_slides, tmp_path, behavior):
    with_libreoffice(monkeypatch)
    fake_powerpoint.behavior = behavior
    notes = []
    files = render(deck_with_two_slides, tmp_path / "out", notes=notes)
    assert [p.name for p in files] == ["slide-01.png", "slide-02.png"] and len(fake_powerpoint.soffice) == 1
    assert len(notes) == 1 and "PowerPoint" in notes[0] and "LibreOffice" in notes[0] and "pptx doctor" in notes[0]
    assert sorted(p.name for p in (tmp_path / "out").iterdir()) == ["slide-01.png", "slide-02.png"]
    assert_empty_export_folder(tmp_path)


def test_a_forced_powerpoint_does_not_fall_back(fake_powerpoint, monkeypatch, deck_with_two_slides, tmp_path):
    with_libreoffice(monkeypatch)
    fake_powerpoint.behavior = "fail"
    with pytest.raises(PptxSkillsError, match="Automation"):
        render(deck_with_two_slides, tmp_path / "out", renderer="powerpoint")
    monkeypatch.setenv("PPTX_SKILLS_RENDERER", "powerpoint")
    with pytest.raises(PptxSkillsError, match="Automation"):
        render(deck_with_two_slides, tmp_path / "out")
    assert fake_powerpoint.soffice == []


def test_cli_render_takes_a_renderer_and_reports_the_fallback(fake_powerpoint, monkeypatch, deck_with_two_slides, tmp_path, capsys):
    with_libreoffice(monkeypatch)
    assert main(["render", str(deck_with_two_slides), "--renderer", "libreoffice", "--out", str(tmp_path / "a")]) == 0
    assert fake_powerpoint.exports == [] and len(fake_powerpoint.soffice) == 1
    capsys.readouterr()
    fake_powerpoint.behavior = "fail"
    assert main(["render", str(deck_with_two_slides), "--out", str(tmp_path / "b"), "--json"]) == 0
    data = json.loads(capsys.readouterr().out)
    assert len(data["files"]) == 2 and "LibreOffice" in data["notes"][0]
    assert main(["render", str(deck_with_two_slides), "--out", str(tmp_path / "c")]) == 0
    assert "LibreOffice" in capsys.readouterr().out.splitlines()[-1]
    assert main(["render", str(deck_with_two_slides), "--renderer", "keynote"]) == 2


def test_contact_sheets_use_the_export_folder_next_to_the_deck(fake_powerpoint, deck_with_two_slides, tmp_path):
    style = tmp_path / "sheets"
    render_layouts(deck_with_two_slides, style / "layouts.png")
    render_examples(deck_with_two_slides, style / "examples.png")
    assert [export["copy"].parent for export in fake_powerpoint.exports] == [tmp_path / "_export"] * 2
    assert sorted(p.name for p in style.iterdir()) == ["examples.png", "layouts.png"]
    assert_empty_export_folder(tmp_path)
    assert hidden_entries(tmp_path) == []


def test_powerpoint_gets_absolute_paths_for_a_relative_deck(fake_powerpoint, monkeypatch, deck_with_two_slides, tmp_path):
    """AppleScript would read a relative path as an HFS path that PowerPoint cannot open."""
    monkeypatch.chdir(tmp_path)
    render(Path(deck_with_two_slides.name), Path("out"))
    (export,) = fake_powerpoint.exports
    assert export["copy"].is_absolute() and export["pdf"].is_absolute()
    assert export["copy"].parent == (tmp_path / "_export").resolve()


def _aged(path, seconds_ago):
    path.write_bytes(b"x")
    stamp = __import__("time").time() - seconds_ago
    os.utime(path, (stamp, stamp))
    return path


def test_old_leftovers_in_the_export_folder_are_swept(fake_powerpoint, deck_with_two_slides, tmp_path):
    """A PDF that PowerPoint wrote after a timed-out run goes on the next run; recent and foreign files stay."""
    export = tmp_path / "_export"
    export.mkdir()
    old = [_aged(export / name, 3600) for name in (
        f"render-{'a' * 32}.pdf", f"render-{'b' * 32}.pptx", f"~$render-{'b' * 32}.pptx", f"~$nder-{'c' * 32}.pptx",
    )]
    recent = _aged(export / f"render-{'d' * 32}.pdf", 60)  # may belong to a render running right now
    foreign = [_aged(export / name, 3600) for name in (
        "notes.pdf", "render-1234.pptx", f"render-{'e' * 32}.pptx.bak", f"my-render-{'f' * 32}.pdf", f"render-{'0123456789ABCDEF' * 2}.pdf",
    )]
    render(deck_with_two_slides, tmp_path / "out")
    assert [p for p in old if p.exists()] == []
    assert recent.exists() and all(p.exists() for p in foreign)


def test_sheets_stop_trying_powerpoint_after_it_failed(fake_powerpoint, monkeypatch, tmp_path, default_pptx, examples_deck, capsys):
    """In auto mode the examples sheet goes straight to LibreOffice once the layouts sheet fell back."""
    with_libreoffice(monkeypatch)
    fake_powerpoint.behavior = "timeout"
    assert main(["init", str(tmp_path / "decks/a"), "--from", str(default_pptx), "--examples", str(examples_deck), "--sheets"]) == 0
    assert len(fake_powerpoint.exports) == 1 and len(fake_powerpoint.soffice) == 2
    out = capsys.readouterr().out
    assert out.count("PowerPoint could not render") == 1
    assert (tmp_path / "decks/a/style/layouts.png").is_file() and (tmp_path / "decks/a/style/examples.png").is_file()


def test_init_takes_a_renderer(fake_powerpoint, monkeypatch, tmp_path, default_pptx):
    with_libreoffice(monkeypatch)
    assert main(["init", str(tmp_path / "decks/a"), "--from", str(default_pptx), "--renderer", "libreoffice", "--sheets"]) == 0
    assert fake_powerpoint.exports == [] and len(fake_powerpoint.soffice) == 1
    deck = load_deck(tmp_path / "decks/a")
    assert main(["init", str(tmp_path / "decks/a"), "--reuse", "--renderer", "libreoffice", "--sheets"]) == 0
    assert fake_powerpoint.exports == [] and deck == load_deck(tmp_path / "decks/a")


def test_one_export_folder_per_project(fake_powerpoint, tmp_path, default_pptx, examples_deck):
    """Every new folder costs the user a "Grant Access" dialog, so all decks and their sheets share decks/_export."""
    decks = tmp_path / "decks"
    init_deck(decks / "a", default_pptx, examples_deck, sheets=True)  # layouts and examples sheets, rendered while staging
    init_deck(decks / "b", default_pptx, sheets=True)
    render(decks / "a/a.pptx")
    assert main(["style", str(decks / "b"), "--refresh"]) == 0
    assert len(fake_powerpoint.exports) == 5
    assert export_folders(fake_powerpoint) == {(decks / "_export").resolve()}
    assert sorted(p.name for p in decks.iterdir()) == ["_export", "a", "b"]
    assert_empty_export_folder(decks)
    assert hidden_entries(decks) == []


def test_style_takes_a_renderer_and_reports_the_fallback(fake_powerpoint, monkeypatch, tmp_path, default_pptx, capsys):
    with_libreoffice(monkeypatch)
    d = tmp_path / "d"
    init_deck(d, default_pptx, sheets=True)
    assert len(fake_powerpoint.exports) == 1 and (d / "style/layouts.png").is_file()
    assert main(["style", str(d), "--refresh", "--renderer", "libreoffice"]) == 0
    assert len(fake_powerpoint.exports) == 1 and len(fake_powerpoint.soffice) == 1
    capsys.readouterr()
    fake_powerpoint.behavior = "fail"
    assert main(["style", str(d), "--refresh", "--json"]) == 0
    notes = json.loads(capsys.readouterr().out)["notes"]
    assert len(notes) == 1 and "LibreOffice" in notes[0]


def test_init_survives_a_failing_powerpoint(fake_powerpoint, tmp_path, default_pptx, capsys):
    fake_powerpoint.behavior = "timeout"
    assert main(["init", str(tmp_path / "d"), "--from", str(default_pptx), "--sheets"]) == 0
    assert "contact sheets skipped: PowerPoint could not export" in capsys.readouterr().out
    assert sorted(p.name for p in (tmp_path / "d/style").iterdir()) == ["brief.json", "brief.md"]
    assert_empty_export_folder(tmp_path)
    assert hidden_entries(tmp_path) == []  # nor is the init staging folder left


def test_init_survives_an_unknown_renderer_in_the_environment(monkeypatch, tmp_path, default_pptx, capsys):
    monkeypatch.setenv("PPTX_SKILLS_RENDERER", "keynote")
    assert main(["init", str(tmp_path / "d"), "--from", str(default_pptx), "--sheets"]) == 0
    assert "contact sheets skipped: PPTX_SKILLS_RENDERER" in capsys.readouterr().out


def test_doctor_names_the_renderer_and_powerpoints_one_time_setup(fake_powerpoint, monkeypatch, capsys):
    report = doctor()
    assert report["renderer"] == "powerpoint" and report["powerpoint"] == str(fake_powerpoint.app)
    assert report["powerpoint_version"] == "16.99" and report["soffice"] is None
    assert main(["doctor"]) == 0  # PowerPoint and poppler are enough
    out = capsys.readouterr().out
    assert "✓ renderer: PowerPoint" in out and "16.99" in out
    assert "Privacy & Security → Automation" in out and "Grant Access" in out and "LibreOffice" in out
    monkeypatch.setenv("PPTX_SKILLS_RENDERER", "libreoffice")
    assert main(["doctor", "--json"]) == 1  # the forced LibreOffice is missing
    data = json.loads(capsys.readouterr().out)
    assert data["renderer"] is None and data["requested_renderer"] == "libreoffice" and data["ok"] is False
    monkeypatch.delenv("PPTX_SKILLS_RENDERER")
    monkeypatch.setattr(render_mod, "find_pdftoppm", lambda: None)
    assert main(["doctor"]) == 1
    assert "✗ pdftoppm" in capsys.readouterr().out


def test_doctor_without_any_renderer(monkeypatch, capsys):
    monkeypatch.setattr(render_mod, "find_soffice", lambda: None)
    monkeypatch.setattr(render_mod, "find_pdftoppm", lambda: Path("/x/pdftoppm"))
    assert doctor()["renderer"] is None
    assert main(["doctor"]) == 1
    out = capsys.readouterr().out
    assert "✗ renderer" in out and "brew install --cask libreoffice" in out


REAL_POWERPOINT = os.environ.get("PPTX_SKILLS_TEST_POWERPOINT") == "1"


def _open_presentations() -> str:
    script = ('if application "Microsoft PowerPoint" is running then\n'
              'tell application "Microsoft PowerPoint" to return name of every presentation\nend if')
    return subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=30).stdout


@pytest.mark.real_powerpoint
@pytest.mark.skipif(not REAL_POWERPOINT, reason="drives the real PowerPoint; set PPTX_SKILLS_TEST_POWERPOINT=1 to run it")
def test_render_real_powerpoint(make_deck):
    if render_mod.find_powerpoint() is None or find_pdftoppm() is None:
        pytest.skip("PowerPoint or poppler not installed")
    before = _open_presentations()
    # PowerPoint writes only into the exact folders the user granted. This folder (git-ignored) holds the deck, so
    # its exchange folder is real-powerpoint-render/_export: one grant ever, as both stay between runs.
    root = Path(__file__).resolve().parent / "real-powerpoint-render"
    root.mkdir(exist_ok=True)
    try:
        deck = Path(shutil.copy(hidden_deck(make_deck), root / "hidden.pptx"))
        out = render(deck, root / "slides", dpi=20, renderer="powerpoint")
        assert [p.name for p in out] == ["slide-01.png", "slide-02.png", "slide-03.png"]
        colors = []
        for path in out:
            with Image.open(path) as image:
                colors.append(image.convert("RGB").getpixel((image.width // 2, image.height // 2)))
        expected = [(255, 0, 0), (0, 255, 0), (0, 0, 255)]  # page n is slide n, hidden or not
        assert all(abs(c - e) <= 8 for got, want in zip(colors, expected) for c, e in zip(got, want)), colors
        assert render_layouts(deck, root / "layouts.png", renderer="powerpoint").is_file()
        assert_empty_export_folder(root)
        assert hidden_entries(root) == []
    finally:
        for item in root.iterdir():
            if item.name != "_export":
                shutil.rmtree(item) if item.is_dir() else item.unlink()
    assert _open_presentations() == before  # the user's presentations stay as they were, no copy is left open
