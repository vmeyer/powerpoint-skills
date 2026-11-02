from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Cm

from conftest import finding, layout_named
from pptx_skills.checks import run_checks as _run_checks


def run_checks(*args, **kwargs):
    """These tests exercise every rule, so they run the strict selection."""
    return _run_checks(*args, strict=True, **kwargs)


def build(make_deck, fill):
    def builder(prs):
        fill(prs.slides.add_slide(layout_named(prs, "Blank")))

    return make_deck(builder)


def box(slide, name, x, y, w, h, text="Text"):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
    shape.name = name
    if text:
        shape.text_frame.text = text
    return shape


def rules(report, rule):
    return [f for f in report["findings"] if f["rule"] == rule]


def test_off_slide_child_in_group(make_deck):
    def fill(slide):
        group = slide.shapes.add_group_shape()
        inner = group.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(30), Cm(2), Cm(6), Cm(2))
        inner.name = "inner"
        inner.text_frame.text = "x"

    f = finding(run_checks(build(make_deck, fill)), "off-slide")
    assert f["severity"] == "error" and f["shape"] == "inner"


def test_overlap_vs_containment(make_deck):
    def fill(slide):
        box(slide, "outer", 2, 2, 10, 6)
        box(slide, "label", 3, 3, 4, 1)
        box(slide, "a", 15, 2, 6, 4)
        box(slide, "b", 18, 4, 6, 4)

    report = run_checks(build(make_deck, fill))
    found = rules(report, "overlap")
    assert len(found) == 1 and found[0]["severity"] == "warning"
    assert "a" in found[0]["message"] and "b" in found[0]["message"]


def test_overlap_ignores_textless_shapes(make_deck):
    def fill(slide):
        box(slide, "a", 2, 2, 6, 4)
        box(slide, "b", 5, 3, 6, 4, text="")

    assert finding(run_checks(build(make_deck, fill)), "overlap") is None


def near_aligned_for(make_deck, left, top=4.0, second_width=7):
    def fill(slide):
        box(slide, "a", 1.50, 1, 5, 2)
        box(slide, "b", left, top, second_width, 3)

    return rules(run_checks(build(make_deck, fill)), "near-aligned")


def test_near_aligned(make_deck):
    found = near_aligned_for(make_deck, 1.79)
    assert len(found) == 1 and found[0]["severity"] == "warning"
    assert near_aligned_for(make_deck, 1.50) == []
    assert near_aligned_for(make_deck, 2.00) == []  # exactly the limit: not flagged
    assert near_aligned_for(make_deck, 2.20) == []


def test_near_aligned_ignores_guide_rounding(make_deck):
    assert near_aligned_for(make_deck, 1.505) == []


def test_near_aligned_needs_shared_column_or_row(make_deck):
    # left edges 0.2 apart but the spans do not overlap horizontally -> different columns
    assert near_aligned_for(make_deck, 6.7, top=4.0, second_width=2) == []
    # tops 0.2 apart but the vertical spans do not overlap -> different rows
    def fill(slide):
        box(slide, "a", 1, 1.0, 4, 2)
        box(slide, "b", 8, 1.2, 4, 2)  # same row: tops differ by 0.2
        box(slide, "c", 8, 9.0, 4, 2)
        box(slide, "d", 20, 9.2, 3, 2)  # same row as c, not as b

    found = rules(run_checks(build(make_deck, fill)), "near-aligned")
    assert {f["shape"] for f in found} == {"b", "d"}
    assert all("top" in f["message"] for f in found)


def test_near_aligned_reports_all_edges_in_one_finding(make_deck):
    found = near_aligned_for(make_deck, 1.70, second_width=5)
    assert len(found) == 1 and "left" in found[0]["message"] and "right" in found[0]["message"]


def test_arrows_and_small_badges_are_not_layout_boxes(make_deck):
    def fill(slide):
        box(slide, "a", 1.50, 1, 5, 2)
        arrow = slide.shapes.add_shape(MSO_SHAPE.UP_ARROW, Cm(1.7), Cm(4), Cm(2), Cm(2))
        arrow.name = "arrow"
        badge = box(slide, "badge", 1.7, 7, 3, 0.6, text="Today")
        badge.name = "badge"
        box(slide, "over", 1.6, 1.2, 4, 0.6, text="inside")

    report = run_checks(build(make_deck, fill))
    assert rules(report, "near-aligned") == [] and rules(report, "overlap") == []


def test_rotated_text_box_uses_rotated_bounds(make_deck):
    def fill(slide):
        axis = box(slide, "axis", -2.46, 8, 9.19, 1.27, text="Axis")
        axis.rotation = 270  # rotated bounds: 1.27 wide, 9.19 high, around x = 2.1
        box(slide, "neighbour", 4, 5, 6, 3, text="Next")

    report = run_checks(build(make_deck, fill))
    assert finding(report, "off-slide") is None and finding(report, "overlap") is None


def test_rotated_shape_that_leaves_the_slide(make_deck):
    def fill(slide):
        s = box(slide, "tilted", 0, 3, 2, 8, text="Tilt")  # on the slide unrotated
        s.rotation = 45  # the 45 degree bounds are 7 cm wide around x = 1

    assert finding(run_checks(build(make_deck, fill)), "off-slide")["shape"] == "tilted"


def test_uneven_gaps(make_deck):
    def fill(slide):
        box(slide, "a", 1, 2, 4, 3)
        box(slide, "b", 5.3, 2, 4, 3)
        box(slide, "c", 9.9, 2, 4, 3)

    assert finding(run_checks(build(make_deck, fill)), "uneven-gaps")["severity"] == "warning"


def test_even_gaps_are_clean(make_deck):
    def fill(slide):
        for i in range(3):
            box(slide, f"s{i}", 1 + i * 4.5, 2, 4, 3)

    assert finding(run_checks(build(make_deck, fill)), "uneven-gaps") is None


def test_rgb_rule_ignores_scheme_aliases(make_deck):
    def fill(slide):
        for i, (theme, bright) in enumerate([(MSO_THEME_COLOR.TEXT_1, 0), (MSO_THEME_COLOR.BACKGROUND_1, -0.1), (MSO_THEME_COLOR.ACCENT_1, 0)]):
            s = box(slide, f"s{i}", 1 + i * 6, 2, 5, 2, text="")
            s.fill.solid()
            s.fill.fore_color.theme_color = theme
            s.fill.fore_color.brightness = bright
        raw = box(slide, "raw", 1, 8, 5, 2, text="")
        raw.fill.solid()
        raw.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)

    found = rules(run_checks(build(make_deck, fill)), "rgb-color")
    assert [f["shape"] for f in found] == ["raw"]


def test_foreign_font(make_deck):
    def fill(slide):
        s = box(slide, "odd", 1, 2, 8, 2, text="hello")
        s.text_frame.paragraphs[0].runs[0].font.name = "Comic Sans MS"
        t = box(slide, "theme", 1, 6, 8, 2, text="hello")
        t.text_frame.paragraphs[0].runs[0].font.name = "+mn-lt"

    found = rules(run_checks(build(make_deck, fill)), "foreign-font")
    assert [f["shape"] for f in found] == ["odd"]


def test_missing_font_once(make_deck, monkeypatch):
    monkeypatch.setattr("pptx_skills.checks.style.find_font", lambda family, *a, **k: None)
    deck = make_deck(lambda prs: [prs.slides.add_slide(layout_named(prs, "Blank")) for _ in range(2)])
    found = rules(run_checks(deck), "missing-font")
    assert len(found) == 1 and found[0]["slide"] is None


def test_alt_text(make_deck, tmp_path):
    from PIL import Image

    img = tmp_path / "p.png"
    Image.new("RGB", (4, 4), "red").save(img)

    def fill(slide):
        pic = slide.shapes.add_picture(str(img), Cm(1), Cm(1), Cm(3), Cm(3))
        pic.name = "pic"
        pic._element.nvPicPr.cNvPr.set("descr", "")

    f = finding(run_checks(build(make_deck, fill)), "alt-text")
    assert f["severity"] == "info" and f["shape"] == "pic"


def test_outside_content_skips_titles(make_deck, tmp_path, monkeypatch):
    from pptx_skills.checks import Context

    def fill(slide):
        box(slide, "wide", 0.2, 5, 5, 2)
        box(slide, "inside", 3, 5, 5, 2)

    deck = build(make_deck, fill)
    monkeypatch.setattr(
        "pptx_skills.checks.load_brief", lambda d: {"content_area": [1.27, 4, 20, 12]}
    )
    monkeypatch.setattr("pptx_skills.checks.find_deck_dir", lambda p: tmp_path)
    monkeypatch.setattr("pptx_skills.checks.load_deck", lambda d: {})
    found = rules(run_checks(deck), "outside-content")
    assert [f["shape"] for f in found] == ["wide"]


def test_uneven_gaps_fix_names_a_real_command(make_deck):
    import re

    from pptx_skills.cli import main

    def fill(slide):
        for i, x in enumerate((2, 8, 15)):
            box(slide, f"b{i}", x, 5, 5, 3)

    fix = finding(run_checks(build(make_deck, fill)), "uneven-gaps")["fix"]
    commands = re.findall(r"`pptx ([a-z-]+)", fix)
    assert commands == ["grid"] and all(main([c, "--help"]) == 0 for c in commands)
