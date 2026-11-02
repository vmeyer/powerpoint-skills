"""Short change sets: names are optional, text in added shapes has sensible defaults, hex colors work everywhere."""

import json

import pytest
from pptx.util import Cm

from conftest import layout_named, png_bytes
from pptx_skills.changeset import apply_changeset, resolve_shape
from pptx_skills.cli import main
from pptx_skills.colors import read_color
from pptx_skills.deckdir import init_deck
from pptx_skills.errors import PptxSkillsError
from pptx_skills.opc import open_presentation
from pptx.oxml.ns import qn


def apply(deck, ops, outside_deck=True):
    return apply_changeset(deck, {"label": "t", "ops": ops}, outside_deck=outside_deck)


@pytest.fixture
def deck(make_deck):
    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        slide.shapes.add_textbox(Cm(20), Cm(15), Cm(2), Cm(1)).name = "rect-1"  # taken: the next rect is rect-2
        prs.slides.add_slide(layout_named(prs, "Blank"))

    return make_deck(build, name="blank.pptx")


def names(deck, number=1):
    return [s.name for s in open_presentation(deck).slides[number - 1].shapes]


def test_names_are_optional_and_unique_per_slide(deck, tmp_path):
    picture = tmp_path / "p.png"
    picture.write_bytes(png_bytes())
    res = apply(deck, [
        {"op": "add_shape", "slide": 1, "shape": "rect", "box": [1, 1, 3, 2]},
        {"op": "add_shape", "slide": 1, "shape": "rect", "box": [5, 1, 3, 2]},
        {"op": "add_shape", "slide": 1, "shape": "textbox", "box": [1, 4, 3, 2], "text": "Hi"},
        {"op": "add_shape", "slide": 1, "name": "own", "shape": "ellipse", "box": [9, 1, 2, 2]},
        {"op": "add_line", "slide": 1, "from": [1, 8], "to": [5, 8], "color": "tx1", "width_pt": 1},
        {"op": "add_table", "slide": 1, "box": [1, 9, 8, 2], "rows": [["a", "b"]]},
        {"op": "add_chart", "slide": 1, "box": [10, 9, 8, 6], "type": "line", "categories": ["x"],
         "series": [{"name": "s", "values": [1]}]},
        {"op": "add_picture", "slide": 1, "path": str(picture), "box": [20, 1, 3, 2], "alt": "A red square"},
        {"op": "group", "slide": 1, "targets": [{"name": "rect-2"}, {"name": "rect-3"}]},
        {"op": "add_shape", "slide": 2, "shape": "rect", "box": [1, 1, 3, 2]},
    ])
    assert names(deck) == ["rect-1", "textbox-1", "own", "line-1", "table-1", "chart-1", "picture-1", "group-1"]
    assert [s.name for s in resolve_shape(open_presentation(deck).slides[0], {"name": "group-1"}).shapes] == ["rect-2", "rect-3"]
    assert names(deck, 2) == ["rect-1"]
    assert res["auto_names"] == {"1": ["rect-2", "rect-3", "textbox-1", "line-1", "table-1", "chart-1", "picture-1", "group-1"],
                                 "2": ["rect-1"]}


def test_given_names_must_still_be_unique(deck):
    with pytest.raises(PptxSkillsError, match="already exists"):
        apply(deck, [{"op": "add_shape", "slide": 1, "name": "rect-1", "shape": "rect", "box": [1, 1, 3, 2]}])


def test_cli_reports_the_given_names(deck, tmp_path, capsys):
    cs = tmp_path / "cs.json"
    cs.write_text(json.dumps({"ops": [{"op": "add_shape", "slide": 2, "shape": "roundRect", "box": [1, 1, 3, 2]}]}))
    assert main(["apply", str(deck), str(cs), "--outside-deck"]) == 0
    assert "named on slide 2: roundRect-1" in capsys.readouterr().out


def _text_props(shape):
    tf = shape.text_frame
    run = tf.paragraphs[0].runs[0]
    return {
        "size": run.font.size.pt,
        "color": read_color(run._r.rPr.find(qn("a:solidFill"))),
        "align": tf.paragraphs[0].alignment,
        "anchor": tf.vertical_anchor,
        "insets": [round(v / 360000, 2) for v in (tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom)],
    }


@pytest.mark.parametrize("kind", ["rect", "textbox"])
def test_text_in_added_shapes_gets_defaults(deck, kind):
    apply(deck, [{"op": "add_shape", "slide": 2, "name": "t", "shape": kind, "box": [1, 1, 8, 3], "text": "Plain text"}])
    props = _text_props(resolve_shape(open_presentation(deck).slides[1], {"name": "t"}))
    assert props == {"size": 14, "color": "tx1", "align": 1, "anchor": 1, "insets": [0.2, 0.2, 0.2, 0.2]}


def test_what_the_op_gives_wins_over_the_defaults(deck):
    apply(deck, [{"op": "add_shape", "slide": 2, "name": "t", "shape": "rect", "box": [1, 1, 8, 3],
                  "font": {"color": "accent2"}, "align": "center", "anchor": "middle", "insets": [0, 0, 0, 0],
                  "paragraphs": [{"runs": [{"t": "Big", "size": 40}]}, {"runs": [{"t": "small"}]}]}])
    shape = resolve_shape(open_presentation(deck).slides[1], {"name": "t"})
    props = _text_props(shape)
    assert props == {"size": 40, "color": "accent2", "align": 2, "anchor": 3, "insets": [0, 0, 0, 0]}
    assert shape.text_frame.paragraphs[1].runs[0].font.size.pt == 14


def test_default_size_is_the_masters_smallest_body_size(tmp_path, default_pptx):
    from pptx_skills.changeset import default_text_size

    init_deck(tmp_path / "decks/m", default_pptx)  # python-pptx's template: body levels 32 down to 20 pt
    deck = tmp_path / "decks/m/m.pptx"
    apply(deck, [{"op": "add_slide", "layout": "Blank"},
                 {"op": "add_shape", "slide": 1, "name": "t", "shape": "rect", "box": [1, 1, 8, 3], "text": "x"}],
          outside_deck=False)
    assert _text_props(resolve_shape(open_presentation(deck).slides[0], {"name": "t"}))["size"] == 20
    style = tmp_path / "style.json"
    style.write_text("{}")
    init_deck(tmp_path / "decks/n", None, style_file=style)  # neutral template: no master, 14 pt
    deck = tmp_path / "decks/n/n.pptx"
    apply(deck, [{"op": "add_slide", "layout": "Blank"},
                 {"op": "add_shape", "slide": 1, "name": "t", "shape": "rect", "box": [1, 1, 8, 3], "text": "x"}],
          outside_deck=False)
    assert _text_props(resolve_shape(open_presentation(deck).slides[0], {"name": "t"}))["size"] == 14
    brief = {"text_styles": {"body_pt": [16, 12, 8]}, "rules": {"min_size_pt": 11}}
    assert default_text_size({"sources": {"master": "m.potx"}}, brief) == 11  # never below the minimum size


def test_hex_colors_work_in_every_op_without_a_flag(deck):
    hexed = "#C0FFEE"
    apply(deck, [
        {"op": "add_shape", "slide": 2, "name": "s", "shape": "rect", "box": [1, 1, 8, 3], "fill": hexed,
         "line": {"color": hexed}, "text": "x", "font": {"color": hexed}},
        {"op": "add_shape", "slide": 2, "name": "p", "shape": "rect", "box": [1, 5, 8, 3],
         "pattern": {"preset": "wdUpDiag", "fg": hexed, "bg": "#FFFFFF"}},
        {"op": "add_line", "slide": 2, "name": "l", "from": [1, 9], "to": [5, 9], "color": hexed, "width_pt": 1},
        {"op": "add_table", "slide": 2, "name": "tb", "box": [10, 1, 8, 2], "rows": [["a"]], "fills": {"0,0": hexed}},
        {"op": "set_text", "slide": 2, "target": {"name": "p"}, "paragraphs": [{"runs": [{"t": "y", "color": hexed}]}]},
        {"op": "update_shape", "slide": 2, "target": {"name": "s"}, "fill": "#123456", "allow_rgb": True},  # still accepted
    ])
    slide = open_presentation(deck).slides[1]
    shape = resolve_shape(slide, {"name": "s"})
    assert str(shape.fill.fore_color.rgb) == "123456" and str(shape.line.color.rgb) == "C0FFEE"
    assert str(shape.text_frame.paragraphs[0].runs[0].font.color.rgb) == "C0FFEE"
    assert str(resolve_shape(slide, {"name": "p"}).text_frame.paragraphs[0].runs[0].font.color.rgb) == "C0FFEE"
    assert str(resolve_shape(slide, {"name": "l"}).line.color.rgb) == "C0FFEE"
    assert str(resolve_shape(slide, {"name": "tb"}).table.cell(0, 0).fill.fore_color.rgb) == "C0FFEE"
    with pytest.raises(PptxSkillsError, match="unknown color"):
        apply(deck, [{"op": "update_shape", "slide": 2, "target": {"name": "s"}, "fill": "#C0FFE"}])
