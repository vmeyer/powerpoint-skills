"""Which rules run: real errors plus two writing rules by default, every rule with --strict."""

import json

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Cm, Pt

from conftest import layout_named
from pptx_skills.checks import DEFAULT_RULE_IDS, RULES, run_checks
from pptx_skills.cli import main


def _box(slide, name, text, x, y, w, h, size=14):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
    shape.name = name
    shape.text_frame.word_wrap = True
    shape.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    shape.text_frame.text = text
    for run in shape.text_frame.paragraphs[0].runs:
        run.font.size = Pt(size)
    return shape


def styled_deck(make_deck):
    """Only matters of taste: a small font, near-aligned boxes and a fixed RGB fill. No error, no writing rule."""

    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        _box(slide, "small", "Fine print", 1, 1, 8, 2, size=8)
        _box(slide, "near", "Almost aligned", 1.1, 4, 8, 2)
        red = _box(slide, "red", "Red fill", 12, 1, 8, 2)
        red.fill.solid()
        red.fill.fore_color.rgb = RGBColor(0xFF, 0, 0)

    return make_deck(build, name="styled.pptx")


def test_default_rules_are_registered_and_cover_errors_and_the_writing_rules():
    registered = {r.rule_id for r in RULES}
    assert DEFAULT_RULE_IDS <= registered
    assert DEFAULT_RULE_IDS == {"reopen", "relationships", "duplicate-ids", "overflow", "off-slide", "leftover-text",
                                "keyword-chain", "too-many-lines"}


def test_default_check_leaves_matters_of_taste_alone(make_deck):
    deck = styled_deck(make_deck)
    assert run_checks(deck)["findings"] == []
    strict = {f["rule"] for f in run_checks(deck, strict=True)["findings"]}
    assert {"min-size", "near-aligned", "rgb-color"} <= strict


def test_default_check_still_finds_errors_and_the_writing_rules(make_deck):
    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        _box(slide, "full", " ".join(["word"] * 80), 1, 1, 5, 1)
        _box(slide, "chain", "Speed · Quality · Scale", 1, 5, 20, 1.5)
        _box(slide, "off", "Off the slide", 30, 1, 10, 2)

    rules = {f["rule"] for f in run_checks(make_deck(build))["findings"]}
    assert {"overflow", "keyword-chain", "off-slide"} <= rules


def test_a_missing_theme_font_is_a_note_by_default_and_a_finding_with_strict(make_deck, monkeypatch):
    monkeypatch.setattr("pptx_skills.checks.style.find_font", lambda family, *a, **k: None)
    deck = make_deck(lambda prs: prs.slides.add_slide(layout_named(prs, "Blank")))
    report = run_checks(deck)
    assert report["findings"] == [] and len(report["notes"]) == 1 and "Calibri" in report["notes"][0]
    strict = run_checks(deck, strict=True)
    assert [f["rule"] for f in strict["findings"]] == ["missing-font"] and not strict.get("notes")


def test_no_slides_means_deck_rules_only(make_deck):
    deck = make_deck(lambda prs: _box(prs.slides.add_slide(layout_named(prs, "Blank")), "full", " ".join(["word"] * 80), 1, 1, 5, 1))
    report = run_checks(deck, slides=[])
    assert report["slides"] == [] and report["findings"] == []


def test_cli_strict_runs_every_rule_and_exits_1_only_on_errors(make_deck, capsys):
    deck = styled_deck(make_deck)
    assert main(["check", str(deck)]) == 0
    assert "Checked slides 1: 0 errors" in capsys.readouterr().out
    assert main(["check", str(deck), "--strict", "--json"]) == 1  # min-size is an error
    found = {f["rule"] for f in json.loads(capsys.readouterr().out)["findings"]}
    assert {"min-size", "near-aligned", "rgb-color"} <= found

    def warnings_only(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        _box(slide, "a", "Almost aligned", 1, 1, 8, 2)
        _box(slide, "b", "with this one", 1.1, 4, 8, 2)

    assert main(["check", str(make_deck(warnings_only, name="w.pptx")), "--strict"]) == 0
    assert "near-aligned" in capsys.readouterr().out
