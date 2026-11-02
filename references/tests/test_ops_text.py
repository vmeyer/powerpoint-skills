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
    assert p0._p.pPr.get("lvl") == "1" and p0._p.pPr.find(qn("a:buChar")) is not None
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
