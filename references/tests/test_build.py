"""`pptx build`: apply a change set or run a script, check and render the touched slides in one command."""

import json
from pathlib import Path

import pytest
from PIL import Image

from pptx_skills.cli import main
from pptx_skills.deckdir import init_deck, load_deck
from pptx_skills.opc import file_hash, open_presentation
from pptx_skills.render import find_pdftoppm, find_soffice

REAL = find_soffice() is not None and find_pdftoppm() is not None


@pytest.fixture
def deck(tmp_path):
    init_deck(tmp_path / "decks/b", None)
    return tmp_path / "decks/b/b.pptx"


@pytest.fixture
def fake_render(monkeypatch):
    """Record the slides rendered; write a small PNG per slide like the real render."""
    calls = []

    def render(deck, out_dir=None, slides=None, dpi=120, renderer=None, notes=None, export_dir=None):
        out_dir = Path(out_dir) if out_dir else Path(deck).parent / "render"
        out_dir.mkdir(parents=True, exist_ok=True)
        calls.append({"slides": list(slides), "dpi": dpi, "renderer": renderer})
        files = []
        for n in slides:
            Image.new("RGB", (160, 90), "white").save(out_dir / f"slide-{n:02d}.png")
            files.append(out_dir / f"slide-{n:02d}.png")
        return files

    monkeypatch.setattr("pptx_skills.build.render", render)
    return calls


def write(tmp_path, name, content):
    path = tmp_path / name
    path.write_text(json.dumps(content) if not isinstance(content, str) else content, encoding="utf-8")
    return path


TWO_SLIDES = {"label": "two", "ops": [
    {"op": "add_slide", "layout": "Title Only"},
    {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "Short change sets build slides fast"},
    {"op": "add_shape", "slide": "$last", "shape": "rect", "box": [2, 6, 12, 3], "fill": "accent1", "text": "One"},
    {"op": "add_slide", "layout": "Title Only"},
    {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "A second slide shows the sheet"},
]}


def test_build_applies_checks_and_renders_the_touched_slides(deck, tmp_path, fake_render, capsys):
    assert main(["build", str(deck), str(write(tmp_path, "c.json", TWO_SLIDES)), "--dpi", "80"]) == 0
    out = capsys.readouterr().out
    folder = deck.parent
    assert "Applied 5 op(s) as change 001" in out and "named on slide 1: rect-1" in out
    assert "Checked slides 1,2: 0 errors" in out
    assert str(folder / "render/slide-01.png") in out and str(folder / "render/slide-02.png") in out
    assert str(folder / "render/sheet.png") in out and (folder / "render/sheet.png").is_file()
    assert fake_render == [{"slides": [1, 2], "dpi": 80, "renderer": None}]
    assert (folder / "check/slide-02.json").is_file() and load_deck(folder)["seq"] == 1


def test_build_runs_a_script_and_renders_one_slide_without_a_sheet(deck, tmp_path, fake_render, capsys):
    script = write(tmp_path, "number.py", "slide = add_slide('Blank')\nbox = slide.shapes.add_textbox(cm(2), cm(2), cm(10), cm(4))\nbox.text_frame.text = '42'\n")
    assert main(["build", str(deck), str(script), "--renderer", "libreoffice", "--json"]) == 0
    data = json.loads(capsys.readouterr().out)
    assert data["seq"] == 1 and data["slides"] == [1] and data["sheet"] is None
    assert data["pngs"] == [str(deck.parent / "render/slide-01.png")] and data["summary"]["error"] == 0
    assert fake_render[0]["renderer"] == "libreoffice"
    assert (deck.parent / "changes/001-number.py").is_file()


def test_build_exits_1_on_errors_and_runs_every_rule_with_strict(deck, tmp_path, fake_render, capsys):
    tiny = {"ops": [{"op": "add_slide", "layout": "Blank"},
                    {"op": "add_shape", "slide": "$last", "shape": "rect", "box": [2, 2, 12, 3], "text": "Tiny", "font": {"size": 7}}]}
    assert main(["build", str(deck), str(write(tmp_path, "t.json", tiny))]) == 0  # min-size runs with --strict only
    capsys.readouterr()
    overflow = {"ops": [{"op": "add_shape", "slide": 1, "shape": "rect", "box": [2, 8, 3, 1], "text": "word " * 60}]}
    assert main(["build", str(deck), str(write(tmp_path, "o.json", overflow))]) == 1
    assert "error overflow" in capsys.readouterr().out
    assert len(fake_render) == 2  # the slide is still rendered, so the agent can look at the problem
    restyle = {"ops": [{"op": "update_shape", "slide": 1, "target": {"name": "rect-1"}, "fill": "#EEEEEE"}]}
    assert main(["build", str(deck), str(write(tmp_path, "r.json", restyle)), "--strict"]) == 1
    out = capsys.readouterr().out
    assert "error min-size" in out and "rgb-color" in out


def test_a_failing_change_renders_nothing_and_exits_2(deck, tmp_path, fake_render, capsys):
    before = file_hash(deck)
    assert main(["build", str(deck), str(write(tmp_path, "bad.json", {"ops": [{"op": "delete_slide", "slide": 3}]}))]) == 2
    assert "op 1 (delete_slide)" in capsys.readouterr().err
    assert main(["build", str(deck), str(write(tmp_path, "bad.py", "raise RuntimeError('no')\n"))]) == 2
    assert "nothing was changed" in capsys.readouterr().err
    assert main(["build", str(deck), str(write(tmp_path, "bad.json", "{nope"))]) == 2
    assert "not valid JSON" in capsys.readouterr().err
    assert file_hash(deck) == before and fake_render == []


def test_without_a_renderer_build_keeps_the_change_reports_it_and_exits_2(deck, tmp_path, monkeypatch, capsys):
    monkeypatch.setattr("pptx_skills.render.find_soffice", lambda: None)
    assert main(["build", str(deck), str(write(tmp_path, "c.json", TWO_SLIDES))]) == 2
    captured = capsys.readouterr()
    assert "Applied 5 op(s) as change 001" in captured.out and "Checked slides 1,2" in captured.out
    assert "error: render failed" in captured.err and "pptx doctor" in captured.err
    assert len(open_presentation(deck).slides) == 2


def test_nothing_to_render_after_a_delete(deck, tmp_path, fake_render, capsys):
    assert main(["build", str(deck), str(write(tmp_path, "c.json", TWO_SLIDES))]) == 0
    capsys.readouterr()
    assert main(["build", str(deck), str(write(tmp_path, "d.json", {"ops": [{"op": "delete_slide", "slide": 2}]}))]) == 0
    assert "nothing to render" in capsys.readouterr().out and len(fake_render) == 1


def test_a_missing_theme_font_is_one_note(deck, tmp_path, fake_render, monkeypatch, capsys):
    monkeypatch.setattr("pptx_skills.checks.style.find_font", lambda family, *a, **k: None)
    assert main(["build", str(deck), str(write(tmp_path, "c.json", TWO_SLIDES))]) == 0
    out = capsys.readouterr().out
    assert out.count("Arial is not installed") == 1 and "missing-font" not in out


@pytest.mark.real_render
@pytest.mark.skipif(not REAL, reason="LibreOffice/poppler not installed")
def test_build_real_with_libreoffice(deck, tmp_path):
    assert main(["build", str(deck), str(write(tmp_path, "c.json", TWO_SLIDES)), "--renderer", "libreoffice", "--dpi", "40"]) == 0
    with Image.open(deck.parent / "render/sheet.png") as sheet:
        assert sheet.width > 500
