"""A deck without a master, its theme colors and fonts, and contact sheets only on request."""

import json

import pytest

from pptx_skills import render as render_mod
from pptx_skills.brief import load_brief
from pptx_skills.cli import main
from pptx_skills.deckdir import init_deck, load_deck
from pptx_skills.errors import PptxSkillsError
from pptx_skills.opc import file_hash, open_presentation


def test_init_without_any_source_starts_on_the_neutral_template(tmp_path, capsys):
    d = tmp_path / "decks/plain"
    assert main(["init", str(d)]) == 0
    out = capsys.readouterr().out
    assert "neutral 16:9 template" in out and "pptx theme" in out
    deck = load_deck(d)
    assert deck["sources"] == {"master": None, "examples": None, "style_file": None}
    prs = open_presentation(d / "plain.pptx")
    assert len(prs.slides) == 0 and round(prs.slide_width / prs.slide_height, 2) == 1.78
    assert load_brief(d)["theme"]["fonts"] == {"major": "Arial", "minor": "Arial"}


@pytest.fixture
def renderer_available(monkeypatch):
    rendered = []

    def fake_layouts(deck, out, *args):
        rendered.append(out.name)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(b"png")

    monkeypatch.setattr(render_mod, "find_soffice", lambda: render_mod.Path("/x/soffice"))
    monkeypatch.setattr(render_mod, "find_pdftoppm", lambda: render_mod.Path("/x/pdftoppm"))
    monkeypatch.setattr(render_mod, "render_layouts", fake_layouts)
    return rendered


def test_contact_sheets_only_on_request(tmp_path, default_pptx, renderer_available):
    d = tmp_path / "decks/a"
    assert main(["init", str(d), "--from", str(default_pptx)]) == 0
    assert renderer_available == [] and not (d / "style/layouts.png").exists()
    assert main(["style", str(d)]) == 0
    assert renderer_available == []
    assert main(["style", str(d), "--sheets"]) == 0
    assert renderer_available == ["layouts.png"] and (d / "style/layouts.png").is_file()
    assert main(["style", str(d), "--sheets"]) == 0  # unchanged sources: nothing to render
    assert main(["style", str(d), "--refresh"]) == 0  # forces the sheets
    assert renderer_available == ["layouts.png", "layouts.png"]
    assert main(["init", str(tmp_path / "decks/b"), "--sheets"]) == 0
    assert len(renderer_available) == 3 and (tmp_path / "decks/b/style/layouts.png").is_file()


def test_theme_sets_colors_and_fonts_with_snapshot_and_log(tmp_path, capsys):
    d = tmp_path / "decks/t"
    init_deck(d, None)
    before = (d / "t.pptx").read_bytes()
    assert main(["theme", str(d), "--color", "accent1=#0b5cad", "--color", "tx1=#222222",
                 "--font", "major=Georgia", "--font", "minor=Verdana", "--json"]) == 0
    result = json.loads(capsys.readouterr().out)
    assert result["seq"] == 1 and result["colors"]["accent1"] == "#0B5CAD" and result["colors"]["dk1"] == "#222222"
    assert result["fonts"] == {"major": "Georgia", "minor": "Verdana"}
    brief = load_brief(d)
    assert brief["theme"]["colors"]["accent1"] == "#0B5CAD" and brief["theme"]["fonts"]["major"] == "Georgia"
    assert (d / "snapshots/001.pptx").read_bytes() == before
    logged = json.loads((d / "changes/001-theme.json").read_text())
    assert logged["theme"] == {"colors": {"accent1": "#0B5CAD", "dk1": "#222222"}, "fonts": {"major": "Georgia", "minor": "Verdana"}}
    assert load_deck(d)["last_hash"] == file_hash(d / "t.pptx")
    assert main(["theme", str(d / "t.pptx")]) == 0  # a working copy works as well as its folder; no options: show
    out = capsys.readouterr().out
    assert "accent1 #0B5CAD" in out and "Georgia" in out and load_deck(d)["seq"] == 1


@pytest.mark.parametrize("args,needle", [
    (["--color", "accent9=#000000"], "accent9"),
    (["--color", "accent1=blue"], "#RRGGBB"),
    (["--color", "accent1"], "role=#RRGGBB"),
    (["--font", "body=Arial"], "major or minor"),
    (["--font", "major="], "major or minor"),
])
def test_theme_rejects_bad_options_and_changes_nothing(tmp_path, capsys, args, needle):
    d = tmp_path / "decks/t"
    init_deck(d, None)
    before = file_hash(d / "t.pptx")
    assert main(["theme", str(d), *args]) == 2
    assert needle in capsys.readouterr().err
    assert file_hash(d / "t.pptx") == before and load_deck(d)["seq"] == 0


def test_theme_writes_only_working_copies(default_pptx, potx_file, capsys):
    for path in (default_pptx, potx_file):
        before = file_hash(path)
        assert main(["theme", str(path), "--color", "accent1=#000000"]) == 2
        err = capsys.readouterr().err
        assert "pptx init" in err and "--outside-deck" not in err
        assert file_hash(path) == before


def test_theme_notes_a_font_that_is_not_installed(tmp_path, capsys, monkeypatch):
    monkeypatch.setattr("pptx_skills.theme.find_font", lambda family, *a, **k: None)
    d = tmp_path / "decks/t"
    init_deck(d, None)
    assert main(["theme", str(d), "--font", "major=Nowhere Sans"]) == 0
    assert "Nowhere Sans is not installed" in capsys.readouterr().out


def test_apply_theme_needs_an_existing_deck(tmp_path):
    from pptx_skills.theme import apply_theme

    with pytest.raises(PptxSkillsError, match="does not exist"):
        apply_theme(tmp_path / "missing.pptx", {"accent1": "#000000"}, {})


def test_theme_says_that_renders_are_stale(tmp_path, capsys):
    d = tmp_path / "decks/t"
    init_deck(d, None)
    assert main(["theme", str(d), "--color", "accent1=#123456"]) == 0
    assert "pptx render" in capsys.readouterr().out


def test_a_new_typeface_drops_the_old_fonts_metrics():
    from lxml import etree
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.oxml.ns import qn

    from pptx_skills.neutral import neutral_presentation, set_theme

    prs = neutral_presentation()
    part = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(part.blob)
    latin = root.find(".//" + qn("a:majorFont")).find(qn("a:latin"))
    latin.attrib.update({"panose": "020F0302020204030204", "pitchFamily": "34", "charset": "0"})
    part._blob = etree.tostring(root)
    set_theme(prs, {}, {"major": "Georgia"})
    latin = etree.fromstring(part.blob).find(".//" + qn("a:majorFont")).find(qn("a:latin"))
    assert dict(latin.attrib) == {"typeface": "Georgia"}
