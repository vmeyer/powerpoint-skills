import pytest
from pptx.oxml.ns import qn
from pptx.util import Cm

from conftest import layout_named
from pptx_skills.changeset import apply_changeset, resolve_shape
from pptx_skills.deckdir import init_deck
from pptx_skills.errors import PptxSkillsError
from pptx_skills.geometry import Box, snap_box
from pptx_skills.inventory import inventory
from pptx_skills.opc import open_presentation


def apply(deck, ops):
    return apply_changeset(deck, {"label": "t", "ops": ops})


def slide1(deck):
    return open_presentation(deck).slides[0]


def shape_named(deck, name):
    return resolve_shape(slide1(deck), {"name": name})


def inv_shape(deck, name):
    def find(shapes):
        for s in shapes:
            if s["name"] == name:
                return s
            hit = find(s.get("children", []))
            if hit:
                return hit

    return find(inventory(open_presentation(deck))["slides"][0]["shapes"])


def exists(deck, name):
    return any(s.name == name for s in _all(slide1(deck).shapes))


def _all(shapes):
    for s in shapes:
        yield s
        if s.shape_type is not None and s.shape_type.name == "GROUP":
            yield from _all(s.shapes)


@pytest.fixture
def deck(make_deck):
    return make_deck(lambda prs: prs.slides.add_slide(layout_named(prs, "Blank")), name="blank.pptx")


@pytest.fixture
def deck_with_guides(tmp_path, guided_master):
    deck_dir = tmp_path / "decks/g"
    init_deck(deck_dir, guided_master)
    path = deck_dir / "g.pptx"
    apply(path, [{"op": "add_slide", "layout": "Blank"}])
    return path


@pytest.fixture
def deck_with_two_boxes(make_deck):
    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        for name, x in (("a", 1), ("b", 6)):
            slide.shapes.add_textbox(Cm(x), Cm(1), Cm(4), Cm(2)).name = name

    return make_deck(build, name="boxes.pptx")


def test_snap_box_moves_edges_within_tolerance():
    box = snap_box(Box(1.53, 5, 8, 2), [1.5, 20.0], [5.02], 0.05)
    assert box == Box(1.5, 5.02, 8.03, 1.98)
    assert snap_box(Box(1.8, 5, 2, 2), [1.5], [], 0.05) == Box(1.8, 5, 2, 2)


def test_add_shape_theme_fill_text_and_snap(deck_with_guides):
    apply(deck_with_guides, [{"op": "add_shape", "slide": 1, "name": "cell-r1", "shape": "rect", "box": [1.53, 5, 8, 2],
                              "fill": "accent1+60", "line": "none", "text": "**Kein Ziel**\nLohnt sich nicht", "font": {"size": 12}}])
    s = shape_named(deck_with_guides, "cell-r1")
    info = inv_shape(deck_with_guides, "cell-r1")
    assert info["box"][0] == 1.49  # the left guide of the guided master
    assert info["fill"] == "accent1+60" and info["line"] == "none"
    runs = s.text_frame.paragraphs[0].runs
    assert runs[0].font.bold and runs[0].font.size.pt == 12
    assert s.text_frame.paragraphs[1].runs[0].font.bold is False


def test_no_snap_without_deck_folder(deck):
    apply(deck, [{"op": "add_shape", "slide": 1, "name": "x", "shape": "ellipse", "box": [1.53, 5, 3, 3]}])
    assert inv_shape(deck, "x")["box"] == [1.53, 5.0, 3.0, 3.0]


def test_pattern_fill(deck):
    apply(deck, [{"op": "add_shape", "slide": 1, "name": "hatch", "shape": "rect", "box": [2, 2, 5, 3],
                  "pattern": {"preset": "wdUpDiag", "fg": "accent6", "bg": "lt1"}}])
    assert inv_shape(deck, "hatch")["fill"] == "pattern:wdUpDiag accent6 lt1"


def test_unknown_pattern_preset(deck):
    with pytest.raises(PptxSkillsError, match="pattern preset"):
        apply(deck, [{"op": "add_shape", "slide": 1, "name": "h", "shape": "rect", "box": [2, 2, 5, 3],
                      "pattern": {"preset": "nope", "fg": "accent6", "bg": "lt1"}}])


def test_rgb_requires_flag(deck):
    ops = [{"op": "add_shape", "slide": 1, "name": "r", "shape": "rect", "box": [2, 2, 5, 3], "fill": "#FF0000"}]
    with pytest.raises(PptxSkillsError, match="allow_rgb"):
        apply(deck, ops)
    apply(deck, [dict(ops[0], allow_rgb=True)])
    assert inv_shape(deck, "r")["fill"] == "#FF0000"


def test_duplicate_name_is_an_error(deck):
    op = {"op": "add_shape", "slide": 1, "name": "dup", "shape": "rect", "box": [2, 2, 5, 3]}
    apply(deck, [op])
    with pytest.raises(PptxSkillsError, match="already exists"):
        apply(deck, [dict(op, box=[8, 2, 3, 3])])


def test_textbox_has_no_fill_and_text_options(deck):
    apply(deck, [{"op": "add_shape", "slide": 1, "name": "tb", "shape": "textbox", "box": [2, 2, 6, 2], "text": "Hi",
                  "align": "center", "anchor": "middle", "insets": [0.1, 0.2, 0.3, 0.4]}])
    s = shape_named(deck, "tb")
    assert s._element.find(qn("p:nvSpPr")).find(qn("p:cNvSpPr")).get("txBox") == "1"
    assert inv_shape(deck, "tb")["fill"] == "none"
    assert s.text_frame.paragraphs[0].alignment == 2 and s.text_frame.vertical_anchor == 3
    assert s.text_frame.margin_left == Cm(0.1) and s.text_frame.margin_bottom == Cm(0.4)


def test_round_rect_radius_and_presets(deck):
    apply(deck, [{"op": "add_shape", "slide": 1, "name": "rr", "shape": "roundRect", "box": [2, 2, 5, 3], "radius": 0.2},
                 {"op": "add_shape", "slide": 1, "name": "hp", "shape": "homePlate", "box": [8, 2, 5, 3]}])
    assert shape_named(deck, "rr").adjustments[0] == pytest.approx(0.2)
    prst = shape_named(deck, "hp")._element.find(".//" + qn("a:prstGeom")).get("prst")
    assert prst == "homePlate"
    with pytest.raises(PptxSkillsError, match="roundRect"):
        apply(deck, [{"op": "update_shape", "slide": 1, "target": {"name": "hp"}, "radius": 0.2}])


def test_line_with_arrowhead_dash_and_snap(deck_with_guides):
    apply(deck_with_guides, [{"op": "add_line", "slide": 1, "name": "path", "from": [10, 8], "to": [14.01, 8.02], "color": "dk1",
                              "width_pt": 3, "dash": "dash", "tail": "triangle", "head": "none"}])
    ln = shape_named(deck_with_guides, "path")._element.find(".//" + qn("a:ln"))
    assert ln.find(qn("a:tailEnd")).get("type") == "triangle" and ln.find(qn("a:headEnd")).get("type") == "none"
    assert ln.get("w") == str(3 * 12700) and ln.find(qn("a:prstDash")).get("val") == "dash"
    assert inv_shape(deck_with_guides, "path")["line"] == {"color": "dk1", "width_pt": 3}


def test_line_endpoint_snaps_to_guide(deck_with_guides):
    apply(deck_with_guides, [{"op": "add_line", "slide": 1, "name": "l", "from": [1.52, 5], "to": [12, 5], "color": "dk1", "width_pt": 1}])
    assert inv_shape(deck_with_guides, "l")["box"][0] == 1.49


def test_update_delete_zorder_group_ungroup(deck_with_two_boxes):
    deck = deck_with_two_boxes
    apply(deck, [{"op": "update_shape", "slide": 1, "target": {"name": "a"}, "box": [2, 2, 3, 3]},
                 {"op": "z_order", "slide": 1, "target": {"name": "a"}, "to": "front"},
                 {"op": "group", "slide": 1, "targets": [{"name": "a"}, {"name": "b"}], "name": "ab"},
                 {"op": "ungroup", "slide": 1, "target": {"name": "ab"}},
                 {"op": "delete_shape", "slide": 1, "target": {"name": "b"}}])
    assert inv_shape(deck, "a")["box"] == [2.0, 2.0, 3.0, 3.0]
    assert not exists(deck, "b") and not exists(deck, "ab")


def test_z_order_positions(deck_with_two_boxes):
    deck = deck_with_two_boxes

    def order():
        return [s.name for s in slide1(deck).shapes]

    assert order() == ["a", "b"]
    apply(deck, [{"op": "z_order", "slide": 1, "target": {"name": "a"}, "to": "front"}])
    assert order() == ["b", "a"]
    apply(deck, [{"op": "z_order", "slide": 1, "target": {"name": "a"}, "to": "backward"}])
    assert order() == ["a", "b"]
    apply(deck, [{"op": "z_order", "slide": 1, "target": {"name": "a"}, "to": "forward"}])
    assert order() == ["b", "a"]
    apply(deck, [{"op": "z_order", "slide": 1, "target": {"name": "a"}, "to": "back"}])
    assert order() == ["a", "b"]


def test_group_encloses_children_exactly(deck_with_two_boxes):
    apply(deck_with_two_boxes, [{"op": "group", "slide": 1, "targets": [{"name": "a"}, {"name": "b"}], "name": "ab"}])
    xfrm = shape_named(deck_with_two_boxes, "ab")._element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
    for a, b in (("a:off", "a:chOff"), ("a:ext", "a:chExt")):
        assert dict(xfrm.find(qn(a)).attrib).values() and list(xfrm.find(qn(a)).attrib.values()) == list(xfrm.find(qn(b)).attrib.values())
    assert inv_shape(deck_with_two_boxes, "ab")["box"] == [1.0, 1.0, 9.0, 2.0]


def test_ungroup_converts_scaled_group_coordinates(deck_with_two_boxes):
    deck = deck_with_two_boxes
    apply(deck, [{"op": "group", "slide": 1, "targets": [{"name": "a"}, {"name": "b"}], "name": "ab"}])
    prs = open_presentation(deck)
    group = resolve_shape(prs.slides[0], {"name": "ab"})
    group.width, group.height = group.width * 2, group.height * 3  # scale: children keep chExt
    prs.save(str(deck))
    expected = inv_shape(deck, "b")["box"]
    apply(deck, [{"op": "ungroup", "slide": 1, "target": {"name": "ab"}}])
    assert inv_shape(deck, "b")["box"] == expected == [11.0, 1.0, 8.0, 6.0]
    assert [s.name for s in slide1(deck).shapes] == ["a", "b"]


def test_group_errors(deck_with_two_boxes):
    with pytest.raises(PptxSkillsError, match="not a group"):
        apply(deck_with_two_boxes, [{"op": "ungroup", "slide": 1, "target": {"name": "a"}}])
    with pytest.raises(PptxSkillsError, match="already exists"):
        apply(deck_with_two_boxes, [{"op": "group", "slide": 1, "targets": [{"name": "a"}, {"name": "b"}], "name": "a"}])


def test_update_style_and_rename(deck):
    apply(deck, [{"op": "add_shape", "slide": 1, "name": "s", "shape": "rect", "box": [2, 2, 5, 3], "fill": "accent1"}])
    apply(deck, [{"op": "update_shape", "slide": 1, "target": {"name": "s"}, "name": "t", "fill": "accent2-25",
                  "line": {"color": "dk1", "width_pt": 2}, "text": "Hello"}])
    info = inv_shape(deck, "t")
    assert info["fill"] == "accent2-25" and info["line"] == {"color": "dk1", "width_pt": 2}
    assert shape_named(deck, "t").text_frame.text == "Hello"


def test_failed_op_leaves_deck_untouched(deck):
    before = deck.read_bytes()
    with pytest.raises(PptxSkillsError, match=r"op 2 \(add_shape\)"):
        apply(deck, [{"op": "add_shape", "slide": 1, "name": "ok", "shape": "rect", "box": [2, 2, 5, 3]},
                     {"op": "add_shape", "slide": 1, "name": "bad", "shape": "rect", "box": [2, 2, 5, 3], "fill": "#00FF00"}])
    assert deck.read_bytes() == before


def test_schema_rejects_text_and_paragraphs_together(deck):
    with pytest.raises(PptxSkillsError, match="add_shape"):
        apply(deck, [{"op": "add_shape", "slide": 1, "name": "x", "shape": "rect", "box": [1, 1, 2, 2],
                      "text": "a", "paragraphs": [{"runs": [{"t": "b"}]}]}])
    with pytest.raises(PptxSkillsError, match="add_shape"):
        apply(deck, [{"op": "add_shape", "slide": 1, "name": "x", "shape": "triangle", "box": [1, 1, 2, 2]}])


def test_unstyled_shape_has_no_style_fill_or_line(deck):
    apply(deck, [{"op": "add_shape", "slide": 1, "name": "plain", "shape": "rect", "box": [2, 2, 5, 3]},
                 {"op": "add_shape", "slide": 1, "name": "styled", "shape": "rect", "box": [8, 2, 5, 3],
                  "fill": "accent2", "line": {"color": "dk1", "width_pt": 1}}])
    el = shape_named(deck, "plain")._element
    assert el.find(qn("p:style")) is None
    info = inv_shape(deck, "plain")
    assert info["fill"] is None and info["line"] is None
    styled = shape_named(deck, "styled")._element
    assert styled.find(qn("p:style")) is None
    assert inv_shape(deck, "styled")["fill"] == "accent2"
    assert inv_shape(deck, "styled")["line"] == {"color": "dk1", "width_pt": 1}


def test_allow_rgb_reaches_text_colors(deck):
    op = {"op": "add_shape", "slide": 1, "name": "t", "shape": "rect", "box": [2, 2, 5, 3], "text": "Hi",
          "font": {"color": "#FF0000"}}
    with pytest.raises(PptxSkillsError, match="allow_rgb"):
        apply(deck, [op])
    apply(deck, [dict(op, allow_rgb=True)])
    run = shape_named(deck, "t").text_frame.paragraphs[0].runs[0]
    assert str(run.font.color.rgb) == "FF0000"


def test_font_only_update_restyles_existing_runs(deck):
    apply(deck, [{"op": "add_shape", "slide": 1, "name": "s", "shape": "rect", "box": [2, 2, 5, 3], "text": "**A** b"}])
    apply(deck, [{"op": "update_shape", "slide": 1, "target": {"name": "s"}, "font": {"size": 20, "color": "accent3"}}])
    runs = shape_named(deck, "s").text_frame.paragraphs[0].runs
    assert [r.font.size.pt for r in runs] == [20, 20]
    assert runs[0].font.bold and not runs[1].font.bold  # bold untouched
    assert runs[0].font.color.theme_color is not None


def test_style_only_update_keeps_word_wrap(deck):
    apply(deck, [{"op": "add_shape", "slide": 1, "name": "s", "shape": "rect", "box": [2, 2, 5, 3], "text": "x"}])
    prs = open_presentation(deck)
    resolve_shape(prs.slides[0], {"name": "s"}).text_frame._txBody.find(qn("a:bodyPr")).set("wrap", "none")
    prs.save(str(deck))
    apply(deck, [{"op": "update_shape", "slide": 1, "target": {"name": "s"}, "fill": "accent1"}])
    assert shape_named(deck, "s").text_frame._txBody.find(qn("a:bodyPr")).get("wrap") == "none"
