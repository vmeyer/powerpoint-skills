import json

import pytest

from pptx_skills.brief import build_brief
from pptx_skills.deckdir import init_deck
from pptx_skills.errors import PptxSkillsError
from pptx_skills.neutral import neutral_presentation
from pptx_skills.opc import open_presentation


@pytest.fixture
def style_json(tmp_path):
    path = tmp_path / "style.json"
    path.write_text(json.dumps({"colors": {"accent1": "#123456"}, "fonts": {"major": "Georgia"}}))
    return path


def test_neutral_is_16_9_with_style():
    prs = neutral_presentation({"colors": {"accent1": "#123456"}, "fonts": {"major": "Georgia", "minor": "Verdana"}})
    b = build_brief(prs)
    assert b["slide_size_cm"] == [33.87, 19.05]
    assert b["theme"]["colors"]["accent1"] == "#123456" and b["theme"]["colors"]["accent2"] == "#2E75B6"
    assert b["theme"]["colors"]["dk1"] == "#000000"
    assert b["theme"]["fonts"] == {"major": "Georgia", "minor": "Verdana"}
    assert b["content_area"] == pytest.approx([1.69, 4.45, 30.48, 12.57], abs=0.02)


def test_neutral_defaults_and_layouts_scaled():
    prs = neutral_presentation()
    assert build_brief(prs)["theme"]["fonts"] == {"major": "Arial", "minor": "Arial"}
    assert prs.slide_master.part.part_related_by(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    ).blob.count(b"sysClr") == 0
    for layout in prs.slide_layouts:
        for ph in layout.placeholders:
            if ph.left is not None:
                assert ph.left + ph.width <= prs.slide_width


def test_init_with_style_file_only(tmp_path, style_json):
    d = tmp_path / "decks/neutral"
    init_deck(d, None, style_file=style_json)
    prs = open_presentation(d / "neutral.pptx")
    assert len(prs.slides) == 0
    brief = json.loads((d / "style/brief.json").read_text())
    assert brief["theme"]["colors"]["accent1"] == "#123456"
    assert brief["theme"]["fonts"]["major"] == "Georgia"


def test_style_file_rules_apply_to_any_source(tmp_path, potx_file):
    style = tmp_path / "s.json"
    style.write_text(json.dumps({"rules": {"min_size_pt": 10}}))
    init_deck(tmp_path / "r", potx_file, style_file=style)
    assert json.loads((tmp_path / "r/style/brief.json").read_text())["rules"]["min_size_pt"] == 10


@pytest.mark.parametrize(
    "data, key",
    [
        ({"colors": {"accent9": "#000000"}}, "accent9"),
        ({"colors": {"accent1": "red"}}, "accent1"),
        ({"fonts": {"mayor": "Arial"}}, "mayor"),
        ({"rules": {"max_bullets": 0}}, "max_bullets"),
        ({"rules": {"bogus": 1}}, "bogus"),
        ({"extra": 1}, "extra"),
        ([], "object"),
    ],
)
def test_invalid_style_file(tmp_path, potx_file, data, key):
    bad = tmp_path / "bad.json"
    bad.write_text(json.dumps(data))
    for source in (None, potx_file):
        with pytest.raises(PptxSkillsError, match=key):
            init_deck(tmp_path / "d", source, style_file=bad)
        assert not (tmp_path / "d").exists()
    assert [p.name for p in tmp_path.iterdir() if p.name.startswith(".")] == []
