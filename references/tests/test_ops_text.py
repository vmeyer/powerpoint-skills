import pytest
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

from conftest import layout_named
from pptx_skills.changeset import apply_changeset, resolve_shape
from pptx_skills.colors import read_color
from pptx_skills.errors import PptxSkillsError
from pptx_skills.opc import open_presentation
from pptx_skills.ops_text import write_paragraphs


def apply(deck, ops):
    return apply_changeset(deck, {"label": "t", "ops": ops})


def slide1(deck):
    return open_presentation(deck).slides[0]


@pytest.fixture
def title_deck(make_deck):
    def build(prs):
        prs.slides.add_slide(layout_named(prs, "Title Only"))

    return make_deck(build, name="title.pptx")


@pytest.fixture
def formatted_box_deck(make_deck):
    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        box = slide.shapes.add_textbox(Cm(2), Cm(2), Cm(8), Cm(2))
        box.name = "box"
        run = box.text_frame.paragraphs[0].add_run()
        run.text = "Old"
        run.font.size = Pt(12)
        run.font.color.theme_color = 5  # ACCENT_1

    return make_deck(build, name="box.pptx")


def test_shorthand_bold(title_deck):
    apply(title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "text": "**Bold** lead\nSecond"}])
    p = slide1(title_deck).shapes.title.text_frame.paragraphs
    assert [r.text for r in p[0].runs] == ["Bold", " lead"] and p[0].runs[0].font.bold and not p[0].runs[1].font.bold
    assert p[1].text == "Second" and len(p) == 2


def test_keeps_formatting_not_specified(formatted_box_deck):
    apply(formatted_box_deck, [{"op": "set_text", "slide": 1, "target": {"name": "box"}, "text": "New"}])
    shape = resolve_shape(slide1(formatted_box_deck), {"name": "box"})
    run = shape.text_frame.paragraphs[0].runs[0]
    assert run.text == "New" and run.font.size.pt == 12
    assert read_color(run._r.rPr.find(qn("a:solidFill"))) == "accent1"


def test_paragraph_runs_with_color_level_bullet_and_break(title_deck):
    apply(title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "paragraphs": [
        {"level": 1, "bullet": True, "runs": [
            {"t": "R4", "color": "accent1", "b": True, "size": 14}, {"t": " Review", "i": True},
            {"br": True}, {"t": "next", "color": "accent2+40"}]},
        {"bullet": False, "runs": [{"t": "plain"}]}]}])
    p0, p1 = slide1(title_deck).shapes.title.text_frame.paragraphs
    assert p0._p.pPr.get("lvl") == "1" and p0._p.pPr.find(qn("a:buNone")) is None
    kids = [c for c in p0._p if c.tag in (qn("a:r"), qn("a:br"))]
    assert [c.tag.split("}")[1] for c in kids] == ["r", "r", "br", "r"]
    r0, r1, _, r3 = kids
    assert r0.find(qn("a:rPr")).get("b") == "1" and r0.find(qn("a:rPr")).get("sz") == "1400"
    assert read_color(r0.find(qn("a:rPr")).find(qn("a:solidFill"))) == "accent1"
    assert r1.find(qn("a:rPr")).get("i") == "1"
    assert read_color(r3.find(qn("a:rPr")).find(qn("a:solidFill"))) == "accent2+40"
    assert p1._p.pPr.find(qn("a:buNone")) is not None


def test_inventory_style_break_and_font(title_deck):
    apply(title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "font": {"size": 20, "bold": True},
                        "paragraphs": [{"runs": [{"t": "a"}, {"t": "\n", "br": True}, {"t": "b", "b": False}]}]}])
    p = slide1(title_deck).shapes.title.text_frame.paragraphs[0]
    assert p._p.find(qn("a:br")) is not None
    runs = p.runs
    assert [r.text for r in runs] == ["a", "b"] and runs[0].font.size.pt == 20
    assert runs[0].font.bold is True and runs[1].font.bold is False


def test_target_kicker_role(kicker_master):
    prs = Presentation(str(kicker_master))
    prs.slides.add_slide(layout_named(prs, "Title Only"))
    prs.save(str(kicker_master))
    apply(kicker_master, [{"op": "set_text", "slide": 1, "target": {"role": "kicker"}, "text": "CLASSIFICATION"}])
    shape = resolve_shape(slide1(kicker_master), {"role": "kicker"})
    assert shape.text_frame.text == "CLASSIFICATION"


def test_records_touched_slide(title_deck):
    res = apply(title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "text": "x"}])
    assert res["touched"] == [slide1(title_deck).slide_id]


def test_errors(title_deck, formatted_box_deck):
    with pytest.raises(PptxSkillsError, match="op 1 .set_text."):
        apply(title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}}])
    with pytest.raises(PptxSkillsError, match="op 1 .set_text."):
        apply(title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "text": "a", "paragraphs": []}])
    with pytest.raises(PptxSkillsError, match="raw color"):
        apply(title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"},
                            "paragraphs": [{"runs": [{"t": "a", "color": "#FF0000"}]}]}])
    with pytest.raises(PptxSkillsError, match="no shape matches"):
        apply(title_deck, [{"op": "set_text", "slide": 1, "target": {"name": "ghost"}, "text": "a"}])


def test_write_paragraphs_is_independent_of_slides():
    prs = Presentation()
    slide = prs.slides.add_slide(layout_named(prs, "Blank"))
    box = slide.shapes.add_textbox(0, 0, Cm(5), Cm(2))
    write_paragraphs(box.text_frame, [{"runs": [{"t": "one"}]}, {"runs": [{"t": "two"}]}])
    assert box.text_frame.text == "one\ntwo"
    write_paragraphs(box.text_frame, [{"runs": [{"t": "only"}]}])
    assert box.text_frame.text == "only"


def test_resolve_shape_by_role_name_id_and_groups(make_deck):
    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Title Only"))
        group = slide.shapes.add_group_shape()
        inner = group.shapes.add_textbox(0, 0, Cm(3), Cm(1))
        inner.name = "inner"
        for i in range(2):
            slide.shapes.add_textbox(0, 0, Cm(3), Cm(1)).name = "twin"

    slide = open_presentation(make_deck(build)).slides[0]
    assert resolve_shape(slide, {"role": "title"}).shape_id == slide.shapes.title.shape_id
    inner = resolve_shape(slide, {"name": "inner"})
    assert resolve_shape(slide, {"id": inner.shape_id}).name == "inner"
    with pytest.raises(PptxSkillsError, match="matches 2 shapes"):
        resolve_shape(slide, {"name": "twin"})
    with pytest.raises(PptxSkillsError, match="no shape matches"):
        resolve_shape(slide, {"name": "ghost"})
    with pytest.raises(PptxSkillsError, match="exactly one of"):
        resolve_shape(slide, {"role": "title", "id": 1})


@pytest.fixture
def hand_formatted_deck(make_deck):
    """First paragraph: bold run, explicit hanging indent and a custom bullet; second: a buNone heading line."""
    from lxml import etree
    from pptx.oxml.ns import nsdecls

    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        box = slide.shapes.add_textbox(Cm(2), Cm(2), Cm(8), Cm(4))
        box.name = "box"
        body = box.text_frame._txBody
        for p in body.findall(qn("a:p")):
            body.remove(p)
        for xml in (
            '<a:p %s><a:pPr marL="342900" indent="-342900" algn="r"><a:buClr><a:schemeClr val="accent2"/></a:buClr>'
            '<a:buChar char="-"/></a:pPr><a:r><a:rPr lang="en-US" b="1"/><a:t>Old</a:t></a:r></a:p>',
            '<a:p %s><a:pPr marL="0" indent="0"><a:buNone/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Head</a:t></a:r></a:p>',
        ):
            body.append(etree.fromstring(xml % nsdecls("a")))

    return make_deck(build, name="hand.pptx")


def _ppr_after(deck, paragraphs):
    apply(deck, [{"op": "set_text", "slide": 1, "target": {"name": "box"}, "paragraphs": paragraphs}])
    shape = resolve_shape(slide1(deck), {"name": "box"})
    return [p._p.pPr for p in shape.text_frame.paragraphs]


def test_neither_level_nor_bullet_keeps_template_ppr(hand_formatted_deck):
    ppr = _ppr_after(hand_formatted_deck, [{"runs": [{"t": "a"}]}, {"runs": [{"t": "b"}]}])[1]
    assert ppr.get("marL") == "342900" and ppr.find(qn("a:buChar")).get("char") == "-"


def test_level_drops_stale_margins(hand_formatted_deck):
    ppr = _ppr_after(hand_formatted_deck, [{"level": 1, "runs": [{"t": "a"}]}])[0]
    assert ppr.get("lvl") == "1" and ppr.get("marL") is None and ppr.get("indent") is None
    assert ppr.get("algn") == "r"


def test_bullet_true_keeps_custom_bullet(hand_formatted_deck):
    ppr = _ppr_after(hand_formatted_deck, [{"bullet": True, "runs": [{"t": "a"}]}])[0]
    assert ppr.find(qn("a:buChar")).get("char") == "-" and ppr.find(qn("a:buClr")) is not None
    assert ppr.get("marL") == "342900"


def test_bullet_true_after_bu_none_template_lets_layout_decide(hand_formatted_deck):
    from pptx_skills.ops_text import write_paragraphs

    prs = open_presentation(hand_formatted_deck)
    box = resolve_shape(prs.slides[0], {"name": "box"})
    body = box.text_frame._txBody
    body.remove(body.findall(qn("a:p"))[0])  # the heading line becomes the template
    write_paragraphs(box.text_frame, [{"bullet": True, "runs": [{"t": "a"}]}])
    ppr = box.text_frame.paragraphs[0]._p.pPr
    assert ppr.find(qn("a:buNone")) is None and ppr.get("marL") is None and ppr.find(qn("a:buChar")) is None


def test_bullet_false_removes_bullet_and_indent(hand_formatted_deck):
    ppr = _ppr_after(hand_formatted_deck, [{"bullet": False, "runs": [{"t": "a"}]}])[0]
    assert ppr.find(qn("a:buNone")) is not None and ppr.find(qn("a:buChar")) is None and ppr.find(qn("a:buClr")) is None
    assert ppr.get("marL") == "0" and ppr.get("indent") == "0"
    assert [c.tag.split("}")[1] for c in ppr] == ["buNone"]


def test_shorthand_bold_is_set_by_markers_only(hand_formatted_deck):
    apply(hand_formatted_deck, [{"op": "set_text", "slide": 1, "target": {"name": "box"}, "text": "**Label** rest"}])
    runs = resolve_shape(slide1(hand_formatted_deck), {"name": "box"}).text_frame.paragraphs[0].runs
    assert [r.text for r in runs] == ["Label", " rest"] and runs[0].font.bold is True and runs[1].font.bold is False


@pytest.fixture
def bold_title_deck(make_deck):
    """A master whose title style is bold, and one "Title Only" slide."""

    def build(prs):
        title_style = prs.slide_master._element.find(qn("p:txStyles")).find(qn("p:titleStyle"))
        title_style.find(qn("a:lvl1pPr")).find(qn("a:defRPr")).set("b", "1")
        prs.slides.add_slide(layout_named(prs, "Title Only"))

    return make_deck(build, name="bold-title.pptx")


def _title_weights(deck):
    return [r._r.find(qn("a:rPr")).get("b") for r in slide1(deck).shapes.title.text_frame.paragraphs[0].runs]


def test_shorthand_without_markers_inherits_the_master_weight(bold_title_deck):
    apply(bold_title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "text": "Three quick wins come first"}])
    assert _title_weights(bold_title_deck) == [None]  # the master's bold title style applies


def test_shorthand_with_markers_writes_every_weight(bold_title_deck):
    apply(bold_title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "text": "**Three wins** come first"}])
    assert _title_weights(bold_title_deck) == ["1", "0"]
    apply(bold_title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "text": "Now plain again"}])
    assert _title_weights(bold_title_deck) == [None]  # the bold copied from the old first run is dropped


def test_shorthand_without_markers_drops_a_template_bold(hand_formatted_deck):
    apply(hand_formatted_deck, [{"op": "set_text", "slide": 1, "target": {"name": "box"}, "text": "plain"}])
    run = resolve_shape(slide1(hand_formatted_deck), {"name": "box"}).text_frame.paragraphs[0].runs[0]
    assert run.text == "plain" and run._r.find(qn("a:rPr")).get("b") is None


def test_font_bold_applies_to_runs_without_markers(hand_formatted_deck):
    apply(hand_formatted_deck, [{"op": "set_text", "slide": 1, "target": {"name": "box"}, "text": "plain", "font": {"bold": False}}])
    assert resolve_shape(slide1(hand_formatted_deck), {"name": "box"}).text_frame.paragraphs[0].runs[0].font.bold is False


def test_raw_color_errors_suggest_allow_rgb_only_where_the_op_takes_it(title_deck):
    run = {"t": "a", "color": "#FF0000"}
    with pytest.raises(PptxSkillsError, match="raw color") as exc:
        apply(title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "paragraphs": [{"runs": [run]}]}])
    assert "allow_rgb" not in str(exc.value)
    with pytest.raises(PptxSkillsError, match="raw color") as exc:
        apply(title_deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "text": "a", "font": {"color": "#FF0000"}}])
    assert "allow_rgb" not in str(exc.value)
    for shape_op in (
        {"op": "add_shape", "slide": 1, "name": "s", "shape": "rect", "box": [1, 1, 3, 2], "fill": "#FF0000"},
        {"op": "add_shape", "slide": 1, "name": "s", "shape": "rect", "box": [1, 1, 3, 2], "paragraphs": [{"runs": [run]}]},
    ):
        with pytest.raises(PptxSkillsError, match='"allow_rgb": true'):
            apply(title_deck, [shape_op])
