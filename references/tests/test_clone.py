"""Cloning, duplicating and replacing slides across decks, with pictures, charts, links and sections."""

import hashlib
import io
import zipfile

import pytest
from lxml import etree
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.util import Cm

from conftest import (
    SLIDE_ZOOM, apply, assert_sections_consistent, layout_named, linked_deck, png_bytes, sectioned_deck, sections,
    slide_links, title_slide,
)
from pptx_skills import slides
from pptx_skills.changeset import apply_changeset
from pptx_skills.checks import run_checks
from pptx_skills.deckdir import load_deck
from pptx_skills.errors import PptxSkillsError
from pptx_skills.opc import file_hash, open_presentation

R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
SVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"


def _parts(prs, prefix):
    return sorted({str(p.partname) for p in prs.part.package.iter_parts() if p.partname.startswith(prefix)})


def _of_type(slide, kind):
    return [s for s in slide.shapes if s.shape_type == kind]


# --- clone from the examples deck -------------------------------------------------------------

def test_clone_with_picture_and_chart(examples_deck, deck_dir):
    deck = deck_dir / "d.pptx"
    res = apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 1}}])
    prs = open_presentation(deck)
    s = prs.slides[0]
    assert res["touched"] == [s.slide_id]
    assert {sh.shape_type for sh in s.shapes} >= {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART}
    chart_shape = _of_type(s, MSO_SHAPE_TYPE.CHART)[0]
    assert chart_shape.chart.plots[0].categories[0] == "A"
    assert chart_shape.chart.chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED
    source = open_presentation(examples_deck).slides[0]
    src_chart = _of_type(source, MSO_SHAPE_TYPE.CHART)[0].chart
    workbook = chart_shape.chart.part.chart_workbook.xlsx_part
    assert workbook is not None and workbook.blob == src_chart.part.chart_workbook.xlsx_part.blob
    picture = _of_type(s, MSO_SHAPE_TYPE.PICTURE)[0]
    assert picture.image.blob == png_bytes()
    assert run_checks(deck, all_slides=True)["summary"]["error"] == 0


def test_clone_twice_shares_the_picture_but_not_the_chart(examples_deck, deck_dir):
    deck = deck_dir / "d.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 1}},
                 {"op": "clone_slide", "from": {"deck": "examples", "slide": 1}}])
    prs = open_presentation(deck)
    assert len(prs.slides) == 2
    assert len(_parts(prs, "/ppt/media/")) == 1
    assert len(_parts(prs, "/ppt/charts/")) == 2 and len(_parts(prs, "/ppt/embeddings/")) == 2
    charts = [_of_type(s, MSO_SHAPE_TYPE.CHART)[0].chart.part for s in prs.slides]
    assert charts[0] is not charts[1]


def test_clone_after_a_slide_and_last_ref(examples_deck, deck_dir):
    deck = deck_dir / "d.pptx"
    apply(deck, [{"op": "add_slide", "layout": "Title Only"}, {"op": "add_slide", "layout": "Blank"},
                 {"op": "clone_slide", "from": {"deck": "examples", "slide": 2}, "after": 1},
                 {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "The clone is the last slide made"}])
    prs = open_presentation(deck)
    assert [s.slide_layout.name for s in prs.slides] == ["Title Only", "Title Only", "Blank"]
    assert prs.slides[1].shapes.title.text == "The clone is the last slide made"


def test_missing_layout_and_ole(deck_dir, deck_with_unknown_layout, deck_with_ole):
    deck = deck_dir / "d.pptx"
    before = file_hash(deck)
    with pytest.raises(PptxSkillsError, match=r"op 1 \(clone_slide\): .*layout 'Special layout' not in working copy; available: .*\"Title Only\""):
        apply(deck, [{"op": "clone_slide", "from": {"deck": str(deck_with_unknown_layout), "slide": 1}}])
    with pytest.raises(PptxSkillsError, match="slide 1 uses oleObject, which cannot be cloned"):
        apply(deck, [{"op": "clone_slide", "from": {"deck": str(deck_with_ole), "slide": 1}}])
    assert file_hash(deck) == before


def test_refused_clone_leaves_the_destination_alone(default_pptx, deck_with_ole):
    dst, src = open_presentation(default_pptx), open_presentation(deck_with_ole)
    with pytest.raises(PptxSkillsError, match="cannot be cloned"):
        slides.clone_slide(src, src.slides[0], dst, None)
    assert len(dst.slides) == 0 and len(dst.slides._sldIdLst) == 0


def test_examples_source_needs_a_deck_folder_with_examples(tmp_path, default_pptx, deck_dir_from_default):
    with pytest.raises(PptxSkillsError, match=r'"examples".*not inside a deck folder'):
        apply(default_pptx, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 1}}])
    with pytest.raises(PptxSkillsError, match=r"has no examples deck.*--examples"):
        apply(deck_dir_from_default / "d.pptx", [{"op": "clone_slide", "from": {"deck": "examples", "slide": 1}}])


def test_bad_source_refs(deck_dir, other_deck, tmp_path):
    deck = deck_dir / "d.pptx"
    with pytest.raises(PptxSkillsError, match=r"from .*other\.pptx: slide 4 does not exist; the deck has 1 slide"):
        apply(deck, [{"op": "clone_slide", "from": {"deck": str(other_deck), "slide": 4}}])
    with pytest.raises(PptxSkillsError, match="does not exist"):
        apply(deck, [{"op": "clone_slide", "from": {"deck": str(tmp_path / "nope.pptx"), "slide": 1}}])
    with pytest.raises(PptxSkillsError, match=r'"\$last"'):
        apply(deck, [{"op": "clone_slide", "from": {"deck": str(other_deck), "slide": "$last"}}])
    with pytest.raises(PptxSkillsError, match=r"op 1 \(clone_slide\): .*from"):
        apply(deck, [{"op": "clone_slide", "after": 1}])
    with pytest.raises(PptxSkillsError, match=r"op 1 \(replace_slide\)"):
        apply(deck, [{"op": "replace_slide", "from": {"deck": "self", "slide": 1}}])


# --- duplicate within the working copy ------------------------------------------------------

def test_duplicate_self(deck_with_two_slides):
    apply(deck_with_two_slides, [{"op": "clone_slide", "from": {"deck": "self", "slide": 1}}])
    prs = open_presentation(deck_with_two_slides)
    assert [s.shapes.title.text for s in prs.slides] == ["First slide", "Second slide", "First slide"]
    first, copy = prs.slides[0], prs.slides[2]
    assert first.slide_id != copy.slide_id
    assert [(s.shape_id, s.name) for s in first.shapes] == [(s.shape_id, s.name) for s in copy.shapes]
    assert first.part is not copy.part


def test_duplicate_of_a_slide_made_in_the_same_change_set(deck_dir_from_default):
    deck = deck_dir_from_default / "d.pptx"
    apply(deck, [{"op": "add_slide", "layout": "Title Only"},
                 {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "Made and duplicated at once"},
                 {"op": "clone_slide", "from": {"deck": "self", "slide": "$last"}}])
    assert [s.shapes.title.text for s in open_presentation(deck).slides] == ["Made and duplicated at once"] * 2


def test_duplicate_keeps_slide_links_and_shares_pictures(make_deck):
    def build(prs):
        target = prs.slides.add_slide(layout_named(prs, "Blank"))
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        slide.shapes.add_picture(io.BytesIO(png_bytes()), Cm(1), Cm(1), Cm(2), Cm(2))
        slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(5), Cm(1), Cm(3), Cm(2)).click_action.target_slide = target

    deck = make_deck(build, name="links.pptx")
    apply(deck, [{"op": "clone_slide", "from": {"deck": "self", "slide": 2}}])
    prs = open_presentation(deck)
    copy = prs.slides[2]
    assert next(s for s in copy.shapes if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE).click_action.target_slide.slide_id == prs.slides[0].slide_id
    assert len(_parts(prs, "/ppt/media/")) == 1


# --- replace ---------------------------------------------------------------------------------

def test_replace_keeps_position_and_id(deck_with_three_slides, other_deck):
    deck = deck_with_three_slides
    old_id = open_presentation(deck).slides[1].slide_id
    res = apply(deck, [{"op": "replace_slide", "slide": 2, "from": {"deck": str(other_deck), "slide": 1}}])
    prs = open_presentation(deck)
    s = prs.slides[1]
    assert s.slide_id == old_id and s.shapes.title.text == "From other deck"
    assert [x.shapes.title.text for x in prs.slides] == ["First slide", "From other deck", "Third slide"]
    assert len(_parts(prs, "/ppt/slides/slide")) == 3  # the old slide part is gone
    assert res["touched"] == [old_id]


def test_replace_keeps_sections_and_custom_shows(make_deck, other_deck):
    path = sectioned_deck(make_deck)
    prs = open_presentation(path)
    rid = prs.slides._sldIdLst[1].rId
    show = etree.fromstring(
        f'<p:custShowLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="{R}">'
        f'<p:custShow name="Short" id="0"><p:sldLst><p:sld r:id="{rid}"/></p:sldLst></p:custShow></p:custShowLst>'
    )
    prs.part._element.find(qn("p:defaultTextStyle")).addprevious(show)
    prs.save(path)
    before = sections(open_presentation(path))

    apply(path, [{"op": "replace_slide", "slide": 2, "from": {"deck": str(other_deck), "slide": 1}}])
    prs = open_presentation(path)
    assert sections(prs) == before
    assert_sections_consistent(prs)
    ref = next(prs.part._element.iter(qn("p:sld")))
    assert prs.part.related_part(ref.get(qn("r:id"))) is prs.slides[1].part
    assert prs.slides[1].shapes.title.text == "From other deck"


def test_clone_joins_the_section_of_the_slide_before(make_deck):
    path = sectioned_deck(make_deck)
    apply(path, [{"op": "clone_slide", "from": {"deck": "self", "slide": 1}, "after": 1},
                 {"op": "clone_slide", "from": {"deck": "self", "slide": 3}}])
    prs = open_presentation(path)
    ids = [s.slide_id for s in prs.slides]
    assert sections(prs) == [ids[:3], ids[3:]]
    assert_sections_consistent(prs)


def _number_text(slide):
    ph = next(p for p in slide.placeholders if p.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER)
    return ph.text_frame.text


def test_slide_numbers_follow_clone_and_replace(deck_dir_from_default, make_deck):
    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Title Only"))
        slides.copy_footer_placeholders(slide, slide.slide_layout)
        for field in slide.shapes._spTree.iter(qn("a:fld")):
            if field.get("type") == "slidenum":
                field.find(qn("a:t")).text = "7"
        slide.shapes.title.text = "A numbered slide from elsewhere"

    numbered = make_deck(build, name="numbered.pptx")
    deck = deck_dir_from_default / "d.pptx"
    apply(deck, [{"op": "add_slide", "layout": "Title Only"},
                 {"op": "clone_slide", "from": {"deck": "self", "slide": 1}},
                 {"op": "clone_slide", "from": {"deck": str(numbered), "slide": 1}, "after": 1}])
    assert [_number_text(s) for s in open_presentation(deck).slides] == ["1", "2", "3"]
    apply(deck, [{"op": "replace_slide", "slide": 1, "from": {"deck": str(numbered), "slide": 1}}])
    assert [_number_text(s) for s in open_presentation(deck).slides] == ["1", "2", "3"]


# --- what a clone carries ---------------------------------------------------------------------

def test_clone_copies_background_master_graphics_and_notes(make_deck, deck_dir_from_default):
    def build(prs):
        slide = title_slide(prs, "Background and notes")
        slide.background.fill.solid()
        slide.background.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_2
        slide._element.set("showMasterSp", "0")
        slide.notes_slide.notes_text_frame.text = "Say this out loud"

    source = make_deck(build, name="bg.pptx")
    deck = deck_dir_from_default / "d.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": str(source), "slide": 1}}])
    slide = open_presentation(deck).slides[0]
    assert slide.background.fill.fore_color.theme_color == MSO_THEME_COLOR.ACCENT_2
    assert slide._element.get("showMasterSp") == "0"
    assert slide.has_notes_slide and slide.notes_slide.notes_text_frame.text == "Say this out loud"


def _links_deck(make_deck):
    """Slide 2: a run linked to a web page, a shape jumping to slide 1, and a text box with PowerPoint tags."""
    def build(prs):
        target = title_slide(prs, "Jump target")
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        box = slide.shapes.add_textbox(Cm(1), Cm(1), Cm(8), Cm(2))
        box.name = "linked"
        box.text_frame.text = "Read more"
        box.text_frame.paragraphs[0].runs[0].hyperlink.address = "https://example.com/more"
        jump = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1), Cm(4), Cm(3), Cm(2))
        jump.name = "jump"
        jump.click_action.target_slide = target
        tags = Part(PackURI("/ppt/tags/tag1.xml"), "application/vnd.openxmlformats-officedocument.presentationml.tags+xml",
                    prs.part.package, b'<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tag name="ADDIN" val="1"/></p:tagLst>')
        rid = slide.part.relate_to(tags, RT.TAGS)
        box._element.nvSpPr.nvPr.append(etree.fromstring(
            f'<p:custDataLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="{R}"><p:tags r:id="{rid}"/></p:custDataLst>'))

    return make_deck(build, name="links.pptx")


def test_links_and_tags_across_decks(make_deck, deck_dir_from_default):
    source = _links_deck(make_deck)
    deck = deck_dir_from_default / "d.pptx"
    res = apply(deck, [{"op": "clone_slide", "from": {"deck": str(source), "slide": 2}}])
    slide = open_presentation(deck).slides[0]
    linked = next(s for s in slide.shapes if s.name == "linked")
    assert linked.text_frame.paragraphs[0].runs[0].hyperlink.address == "https://example.com/more"
    jump = next(s for s in slide.shapes if s.name == "jump")
    assert jump._element.find(".//" + qn("a:hlinkClick")) is None
    assert slide._element.find(".//" + qn("p:custDataLst")) is None
    assert RT.TAGS not in {r.reltype for r in slide.part.rels.values()}
    assert any("tags" in m for m in res["messages"]) and any("link to another slide" in m for m in res["messages"])
    report = run_checks(deck, all_slides=True)
    assert not [f for f in report["findings"] if f["category"] == "structure"]


def test_svg_pictures_are_copied(make_deck, deck_dir_from_default):
    svg = b'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>'

    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        pic = slide.shapes.add_picture(io.BytesIO(png_bytes()), Cm(1), Cm(1), Cm(2), Cm(2))
        part = Part(PackURI("/ppt/media/image9.svg"), "image/svg+xml", prs.part.package, svg)
        rid = slide.part.relate_to(part, RT.IMAGE)
        pic._element.blipFill.find(qn("a:blip")).append(etree.fromstring(
            f'<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}">'
            f'<asvg:svgBlip xmlns:asvg="{SVG}" xmlns:r="{R}" r:embed="{rid}"/></a:ext></a:extLst>'))

    source = make_deck(build, name="svg.pptx")
    deck = deck_dir_from_default / "d.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": str(source), "slide": 1}},
                 {"op": "clone_slide", "from": {"deck": str(source), "slide": 1}}])
    prs = open_presentation(deck)
    for slide in prs.slides:
        rid = next(slide._element.iter(f"{{{SVG}}}svgBlip")).get(f"{{{R}}}embed")
        part = slide.part.related_part(rid)
        assert part.content_type == "image/svg+xml" and part.blob == svg
    assert len(_parts(prs, "/ppt/media/")) == 2  # one PNG, one SVG, shared by both clones
    assert run_checks(deck, all_slides=True)["summary"]["error"] == 0


def test_prefixes_named_in_attributes_stay_declared(make_deck, deck_dir_from_default, tmp_path):
    """`mc:Choice Requires="a14"` names a prefix only in an attribute value; the copy must still declare it."""
    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        old = slide.part._element
        root = etree.Element(old.tag, nsmap={**old.nsmap, "a14": A14})
        root.extend(list(old))
        slide.part._element = root
        root.find(qn("p:cSld")).find(qn("p:spTree")).append(etree.fromstring(
            f'<mc:AlternateContent xmlns:mc="{MC}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><mc:Choice Requires="a14">'
            '<p:sp><p:nvSpPr><p:cNvPr id="10" name="choice"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></mc:Choice>'
            '<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="10" name="fallback"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>'
            '</mc:Fallback></mc:AlternateContent>'))

    source = make_deck(build, name="mc.pptx")
    deck = deck_dir_from_default / "d.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": str(source), "slide": 1}}])
    name = open_presentation(deck).slides[0].part.partname.lstrip("/")
    with zipfile.ZipFile(deck) as z:
        root = etree.fromstring(z.read(name))
    choice = next(root.iter(f"{{{MC}}}Choice"))
    assert choice.nsmap.get("a14") == A14


# --- cloned example texts (R3) ----------------------------------------------------------------

def _leftovers(deck):
    return {f["shape_id"] for f in run_checks(deck, all_slides=True)["findings"] if f["rule"] == "leftover-text"}


def test_cloned_example_texts_are_recorded_and_flagged(deck_dir):
    deck = deck_dir / "d.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 2}}])
    slide = open_presentation(deck).slides[0]
    title = slide.shapes.title
    box = next(s for s in slide.shapes if s.name == "Example box")
    assert load_deck(deck_dir)["cloned_texts"] == {
        str(slide.slide_id): {str(title.shape_id): "Example title of the building block", str(box.shape_id): "Example label\nsecond line"}
    }  # date, footer and slide number are the master's, not example text
    assert _leftovers(deck) == {title.shape_id, box.shape_id}
    apply(deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "text": "Our own title states the point"}])
    assert _leftovers(deck) == {box.shape_id}


def test_replace_from_the_users_deck_drops_cloned_texts(deck_dir, other_deck):
    deck = deck_dir / "d.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 2}}])
    slide_id = open_presentation(deck).slides[0].slide_id
    assert str(slide_id) in load_deck(deck_dir)["cloned_texts"]
    apply(deck, [{"op": "replace_slide", "slide": 1, "from": {"deck": str(other_deck), "slide": 1}}])
    assert "cloned_texts" not in load_deck(deck_dir)
    assert not _leftovers(deck)
    apply(deck, [{"op": "replace_slide", "slide": 1, "from": {"deck": "examples", "slide": 2}}])
    assert list(load_deck(deck_dir)["cloned_texts"]) == [str(slide_id)]


def test_clone_from_another_deck_records_nothing(deck_dir, other_deck):
    deck = deck_dir / "d.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": str(other_deck), "slide": 1}}])
    assert "cloned_texts" not in load_deck(deck_dir)


def test_examples_by_path_count_as_examples(deck_dir, examples_deck):
    deck = deck_dir / "d.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": str(examples_deck), "slide": 2}}])
    assert len(load_deck(deck_dir)["cloned_texts"]) == 1


def test_duplicate_carries_cloned_texts(deck_dir):
    deck = deck_dir / "d.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 2}},
                 {"op": "clone_slide", "from": {"deck": "self", "slide": "$last"}}])
    prs = open_presentation(deck)
    texts = load_deck(deck_dir)["cloned_texts"]
    assert set(texts) == {str(s.slide_id) for s in prs.slides}
    assert texts[str(prs.slides[0].slide_id)] == texts[str(prs.slides[1].slide_id)]


def test_cloned_texts_follow_deleted_and_reused_slide_ids(deck_dir):
    deck = deck_dir / "d.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 2}}])
    old_id = open_presentation(deck).slides[0].slide_id
    apply(deck, [{"op": "delete_slide", "slide": 1}, {"op": "add_slide", "layout": "Title Only"}])
    assert open_presentation(deck).slides[0].slide_id == old_id + 1  # the freed id is never reused (R26)
    assert "cloned_texts" not in load_deck(deck_dir)
    apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 2}}])
    apply(deck, [{"op": "delete_slide", "slide": 2}])
    assert "cloned_texts" not in load_deck(deck_dir)


def test_failed_apply_records_no_cloned_texts(deck_dir):
    deck = deck_dir / "d.pptx"
    before = (deck_dir / "deck.json").read_bytes()
    with pytest.raises(PptxSkillsError, match=r"op 2 \(delete_slide\)"):
        apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 2}}, {"op": "delete_slide", "slide": 9}])
    assert (deck_dir / "deck.json").read_bytes() == before


def test_cloned_texts_rolled_back_when_the_log_fails(deck_dir, monkeypatch):
    before = (deck_dir / "deck.json").read_bytes()

    def boom(*a, **k):
        raise OSError("disk says no")

    monkeypatch.setattr("pptx_skills.deckdir._save_deck", boom)
    with pytest.raises(PptxSkillsError, match="nothing was changed"):
        apply(deck_dir / "d.pptx", [{"op": "clone_slide", "from": {"deck": "examples", "slide": 2}}])
    assert (deck_dir / "deck.json").read_bytes() == before


def test_outside_a_deck_folder_clone_writes_no_deck_json(tmp_path, default_pptx, examples_deck):
    res = apply_changeset(default_pptx, {"ops": [{"op": "clone_slide", "from": {"deck": str(examples_deck), "slide": 2}}]})
    assert res["seq"] is None and not (tmp_path / "deck.json").exists()
    assert len(open_presentation(default_pptx).slides) == 1


def test_invalid_cloned_texts_in_deck_json(deck_dir):
    import json

    path = deck_dir / "deck.json"
    path.write_text(json.dumps({**json.loads(path.read_text()), "cloned_texts": []}))
    with pytest.raises(PptxSkillsError, match="cloned_texts"):
        load_deck(deck_dir)


def test_image_bytes_are_shared_by_hash(examples_deck, deck_dir):
    deck = deck_dir / "d.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 1}}])
    apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 1}}])
    prs = open_presentation(deck)
    blobs = [hashlib.sha1(p.blob).hexdigest() for p in prs.part.package.iter_parts() if p.partname.startswith("/ppt/media/")]
    assert len(blobs) == 1


def test_replace_then_add_keeps_part_names_unique(deck_with_three_slides, other_deck):
    apply(deck_with_three_slides, [{"op": "replace_slide", "slide": 1, "from": {"deck": str(other_deck), "slide": 1}},
                                   {"op": "clone_slide", "from": {"deck": "self", "slide": 2}},
                                   {"op": "add_slide", "layout": "Blank"}])
    names = zipfile.ZipFile(deck_with_three_slides).namelist()
    assert len(names) == len(set(names))
    prs = open_presentation(deck_with_three_slides)
    assert [s.part.partname for s in prs.slides] == [f"/ppt/slides/slide{n}.xml" for n in range(1, 6)]
    assert [s.slide_layout.name for s in prs.slides] == ["Title Only"] * 4 + ["Blank"]



# --- links into replaced slides, zooms, chart data -------------------------------------------

def _zoom_ids(slide):
    return [el.get("sldId") for el in slide._element.iter(f"{{{SLIDE_ZOOM}}}sldZmObj")]


def _click_target(slide, name):
    el = next(el for el in slide._element.iter(qn("p:cNvPr")) if el.get("name") == name)
    return slide.part.related_part(el.find(qn("a:hlinkClick")).get(qn("r:id")))


def test_replace_repoints_links_to_the_replaced_slide(make_deck, other_deck):
    deck = linked_deck(make_deck, target=2, alternate=True, outline=2)
    old_id = open_presentation(deck).slides[1].slide_id
    apply(deck, [{"op": "replace_slide", "slide": 2, "from": {"deck": str(other_deck), "slide": 1}}])
    names = zipfile.ZipFile(deck).namelist()
    assert len(names) == len(set(names))
    prs = open_presentation(deck)
    first, new = prs.slides[0], prs.slides[1]
    assert new.slide_id == old_id and new.shapes.title.text == "From other deck"
    assert _click_target(first, "jump") is new.part
    assert _click_target(first, "choice-link") is new.part  # a hyperlink inside mc:Choice too
    assert slide_links(prs) == {"/ppt/slides/slide1.xml": ["/ppt/slides/slide2.xml"], "/ppt/viewProps.xml": ["/ppt/slides/slide2.xml"]}
    view = next(p for p in prs.part.package.iter_parts() if p.partname == "/ppt/viewProps.xml")
    entries = list(etree.fromstring(view.blob).iter(qn("p:sld")))
    assert len(entries) == 2 and all(view.related_part(e.get(qn("r:id"))) is new.part for e in entries)


def test_slide_zoom_is_cloned_only_within_the_deck(make_deck, deck_dir_from_default):
    source = linked_deck(make_deck, target=2, zoom=True)
    deck = deck_dir_from_default / "d.pptx"
    before = file_hash(deck)
    with pytest.raises(PptxSkillsError, match="slide 1 uses a Slide Zoom, which cannot be cloned"):
        apply(deck, [{"op": "clone_slide", "from": {"deck": str(source), "slide": 1}}])
    assert file_hash(deck) == before
    apply(source, [{"op": "clone_slide", "from": {"deck": "self", "slide": 1}}])
    prs = open_presentation(source)
    copy, target = prs.slides[3], prs.slides[1]
    assert _zoom_ids(copy) == [str(target.slide_id)]
    assert _click_target(copy, "zoom-fallback") is target.part


@pytest.mark.parametrize("kind, name", [("sectionzoom", "Section Zoom"), ("summaryzoom", "Summary Zoom")])
def test_zooms_are_found_by_their_graphic_data(make_deck, deck_dir_from_default, kind, name):
    uri = f"http://schemas.microsoft.com/office/powerpoint/2016/{kind}"

    def build(prs):
        prs.slides.add_slide(layout_named(prs, "Blank")).shapes._spTree.append(etree.fromstring(
            f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="a14" xmlns:a14="{A14}">'
            '<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvGraphicFramePr><p:cNvPr id="9" name="zoom"/>'
            '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm>'
            f'<a:graphic><a:graphicData uri="{uri}"/></a:graphic></p:graphicFrame></mc:Choice></mc:AlternateContent>'))

    source = make_deck(build, name="zoom.pptx")
    with pytest.raises(PptxSkillsError, match=f"slide 1 uses a {name}, which cannot be cloned"):
        apply(deck_dir_from_default / "d.pptx", [{"op": "clone_slide", "from": {"deck": str(source), "slide": 1}}])


def test_zoom_target_cannot_be_replaced(make_deck, other_deck):
    deck = linked_deck(make_deck, target=2, zoom=True)
    before = file_hash(deck)
    with pytest.raises(PptxSkillsError, match="cannot replace slide 2: it is the target of a Slide Zoom on slide 1; remove the zoom in PowerPoint first"):
        apply(deck, [{"op": "replace_slide", "slide": 2, "from": {"deck": str(other_deck), "slide": 1}}])
    assert file_hash(deck) == before
    apply(deck, [{"op": "replace_slide", "slide": 3, "from": {"deck": str(other_deck), "slide": 1}}])  # not a target
    assert open_presentation(deck).slides[2].shapes.title.text == "From other deck"


@pytest.mark.parametrize("partname, content_type, reltype", [
    ("/ppt/embeddings/Book1.xlsm", "application/vnd.ms-excel.sheet.macroEnabled.12", RT.PACKAGE),
    ("/ppt/embeddings/Book1.xlsb", "application/vnd.ms-excel.sheet.binary.macroEnabled.main", RT.PACKAGE),
    ("/ppt/embeddings/oleObject1.bin", "application/vnd.openxmlformats-officedocument.oleObject", RT.OLE_OBJECT),
])
def test_chart_data_must_be_an_embedded_xlsx(make_deck, deck_dir_from_default, partname, content_type, reltype):
    from pptx.chart.data import CategoryChartData

    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        data = CategoryChartData()
        data.categories = ["A"]
        data.add_series("S", (1,))
        chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Cm(2), Cm(2), Cm(10), Cm(6), data).chart
        ext = chart.part._element.find(qn("c:externalData"))
        old = ext.get(qn("r:id"))
        ext.set(qn("r:id"), chart.part.relate_to(Part(PackURI(partname), content_type, prs.part.package, b"data"), reltype))
        chart.part.drop_rel(old)

    source = make_deck(build, name="workbook.pptx")
    deck = deck_dir_from_default / "d.pptx"
    with pytest.raises(PptxSkillsError, match=r"slide 1 uses a chart whose data is not an embedded \.xlsx workbook, which cannot be cloned"):
        apply(deck, [{"op": "clone_slide", "from": {"deck": str(source), "slide": 1}}])
