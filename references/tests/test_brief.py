import json

import pytest
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Cm

from conftest import layout_named
from pptx_skills.brief import DEFAULT_RULES, build_brief, load_brief, write_brief
from pptx_skills.opc import open_presentation


def test_theme_and_size(default_pptx):
    b = build_brief(Presentation(default_pptx))
    assert b["slide_size_cm"] == [25.4, 19.05]
    assert b["theme"]["fonts"] == {"major": "Calibri", "minor": "Calibri"}
    assert b["theme"]["colors"]["accent1"] == "#4F81BD" and b["theme"]["colors"]["dk1"] == "#000000"
    assert b["text_styles"]["title_pt"] == 44 and b["text_styles"]["body_pt"][:2] == [32, 28]


def test_content_area_from_guides(guided_master):
    b = build_brief(open_presentation(guided_master))
    assert b["content_area_source"] == "guides"
    assert b["content_area"] == pytest.approx([1.5, 4.82, 30.87, 11.88], abs=0.02)


def test_content_area_from_layouts(default_pptx):
    b = build_brief(Presentation(default_pptx))
    assert (b["content_area_source"], b["content_area"]) == ("layouts", [1.27, 4.45, 22.86, 12.57])


def test_three_guides_fall_back_to_layouts(make_deck):
    from lxml import etree

    def build(prs):
        ns = "http://schemas.microsoft.com/office/powerpoint/2012/main"
        ext = etree.fromstring(
            '<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:ext uri="{27BBF7A9-308A-43DC-89C8-2F10F3537804}">'
            f'<p15:sldGuideLst xmlns:p15="{ns}"><p15:guide id="1" pos="100"/><p15:guide id="2" pos="200"/>'
            '<p15:guide id="3" pos="300"/><p15:guide id="4" orient="horz" pos="100"/></p15:sldGuideLst></p:ext></p:extLst>'
        )
        prs.slide_master._element.append(ext)

    b = build_brief(open_presentation(make_deck(build)))
    assert b["content_area_source"] == "layouts" and len(b["guides"]["vertical"]) == 3


def test_layouts_list_placeholders(default_pptx):
    layouts = {x["name"]: x for x in build_brief(Presentation(default_pptx))["layouts"]}
    ph = layouts["Title and Content"]["placeholders"][1]
    assert (ph["idx"], ph["type"], ph["role"], ph["box"]) == (1, "object", "body", [1.27, 4.45, 22.86, 12.57])
    assert ph["prompt"]


def test_rules_default_and_override(default_pptx):
    r = build_brief(Presentation(default_pptx), rules_override={"min_size_pt": 10})["rules"]
    assert r == {**DEFAULT_RULES, "min_size_pt": 10}
    assert DEFAULT_RULES == {
        "min_size_pt": 11, "min_size_small_pt": 9, "max_lines": 3, "max_title_lines": 2,
        "max_bullets": 5, "near_align_cm": 0.3, "snap_cm": 0.05, "title_min_words": 4,
    }


def test_color_usage_from_examples(default_pptx, examples_with_three_accent1_rects):
    b = build_brief(Presentation(default_pptx), Presentation(examples_with_three_accent1_rects))
    assert b["color_usage"]["fills"]["accent1"] == 3
    ex = b["examples"][0]
    assert (ex["number"], ex["layout"], ex["shape_count"], ex["fills"]) == (1, "Blank", 3, {"accent1": 3})
    assert b["typical_gap_cm"] == {"horizontal": 1.0, "vertical": None}


def test_radii_text_colors_and_vertical_gap(default_pptx, make_deck):
    def build(prs):
        s = prs.slides.add_slide(layout_named(prs, "Blank"))
        a = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(1), Cm(1), Cm(6), Cm(2))
        b = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1), Cm(4), Cm(6), Cm(2))
        run = a.text_frame.paragraphs[0].add_run()
        run.text = "x"
        run.font.size = Cm(0.5)
        run.font.color.theme_color = 5
        b.fill.background()

    b = build_brief(Presentation(default_pptx), Presentation(make_deck(build, name="ex.pptx")))
    assert b["corner_radii"] == [0.33]
    assert b["typical_gap_cm"] == {"horizontal": None, "vertical": 1.0}
    assert b["color_usage"]["text"] == {"accent1": 1} and b["examples"][0]["sizes_pt"] == [14.17]
    assert b["color_usage"]["fills"] == {}


def test_write_and_load(tmp_path, default_pptx):
    assert load_brief(tmp_path) is None
    b = build_brief(Presentation(default_pptx))
    write_brief(b, tmp_path / "style")
    assert load_brief(tmp_path) == json.loads(json.dumps(b))
    assert (tmp_path / "style" / "brief.md").read_text().startswith("# Style brief")


def test_agent_section_survives_rewrite(tmp_path, default_pptx):
    b = build_brief(Presentation(default_pptx)); write_brief(b, tmp_path)
    md = (tmp_path / "brief.md").read_text().replace("<!-- /agent:building-blocks -->", "example 3: KPI tiles\n<!-- /agent:building-blocks -->")
    (tmp_path / "brief.md").write_text(md); write_brief(b, tmp_path)
    assert "example 3: KPI tiles" in (tmp_path / "brief.md").read_text()


def test_group_children_and_non_color_fills(default_pptx, make_deck):
    from lxml import etree
    from pptx.enum.dml import MSO_PATTERN_TYPE
    from pptx.oxml.ns import qn

    def build(prs):
        s = prs.slides.add_slide(layout_named(prs, "Blank"))
        group = s.shapes.add_group_shape()
        child = group.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(1), Cm(1), Cm(6), Cm(2))
        child.fill.solid()
        child.fill.fore_color.theme_color = 6  # ACCENT_2
        pattern = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(10), Cm(1), Cm(3), Cm(2))
        pattern.fill.patterned()
        pattern.fill.pattern = MSO_PATTERN_TYPE.DIAGONAL_BRICK
        gradient = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(14), Cm(1), Cm(3), Cm(2))
        gradient.fill.gradient()
        image = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(18), Cm(1), Cm(3), Cm(2))
        sp_pr = image._element.spPr
        for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:pattFill"):
            for el in sp_pr.findall(qn(tag)):
                sp_pr.remove(el)
        blip = etree.SubElement(sp_pr, qn("a:blipFill"))
        etree.SubElement(blip, qn("a:blip"))
        sp_pr.find(qn("a:prstGeom")).addnext(blip)

    b = build_brief(Presentation(default_pptx), Presentation(make_deck(build, name="grp.pptx")))
    assert b["color_usage"]["fills"] == {"accent2": 1}
    assert b["examples"][0]["fills"] == {"accent2": 1}
    assert b["examples"][0]["shape_count"] == 5
    assert b["corner_radii"] == [0.33]


def test_group_children_gap(default_pptx, make_deck):
    def build(prs):
        s = prs.slides.add_slide(layout_named(prs, "Blank"))
        group = s.shapes.add_group_shape()
        group.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1), Cm(1), Cm(4), Cm(2))
        group.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(6), Cm(1), Cm(4), Cm(2))

    b = build_brief(Presentation(default_pptx), Presentation(make_deck(build, name="grp2.pptx")))
    assert b["typical_gap_cm"]["horizontal"] == 1.0


def test_load_brief_rejects_corrupt_json(tmp_path):
    from pptx_skills.errors import PptxSkillsError

    (tmp_path / "style").mkdir()
    (tmp_path / "style" / "brief.json").write_text("{not json")
    with pytest.raises(PptxSkillsError, match="not valid JSON"):
        load_brief(tmp_path)
    (tmp_path / "style" / "brief.json").write_bytes(b"\xff\xfe\x00")
    with pytest.raises(PptxSkillsError, match="not valid JSON"):
        load_brief(tmp_path)


P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main"


def _guide_ext(kind: str, guides: list[tuple[str | None, int]]) -> str:
    """A `p:ext` with slide guides (`SLIDE_GUIDES`) or notes guides (`NOTES_GUIDES`)."""
    tag = "notesGuideLst" if kind == NOTES_GUIDES else "sldGuideLst"
    items = "".join(f'<p15:guide id="{i}"{f" orient=\"{o}\"" if o else ""} pos="{pos}"/>' for i, (o, pos) in enumerate(guides, 1))
    return f'<p:ext uri="{{{kind}}}"><p15:{tag} xmlns:p15="{P15}">{items}</p15:{tag}></p:ext>'


SLIDE_GUIDES, NOTES_GUIDES = "27BBF7A9-308A-43DC-89C8-2F10F3537804", "2D200454-40CA-4A62-9FC3-DE9A4176ACB9"
VERTICAL, HORIZONTAL = [(None, 339), (None, 7340)], [("horz", 1094), ("horz", 3789)]


def _add_ext(root, ext_xml: str) -> None:
    from lxml import etree

    ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    ext_lst = root.find(f"{{{ns}}}extLst")
    if ext_lst is None:
        ext_lst = etree.SubElement(root, f"{{{ns}}}extLst")
    ext_lst.append(etree.fromstring(f'<p:ext xmlns:p="{ns}"' + ext_xml[len("<p:ext"):]))


def test_guides_ignore_notes_guides_and_duplicates(make_deck):
    def build(prs):
        prs.slide_width = Cm(33.87)
        _add_ext(prs.slide_master._element, _guide_ext(SLIDE_GUIDES, VERTICAL + VERTICAL[:1] + HORIZONTAL + HORIZONTAL[1:]))
        _add_ext(prs.part._element, _guide_ext(NOTES_GUIDES, [(None, 500), ("horz", 700)]))
        _add_ext(prs.part._element, _guide_ext(SLIDE_GUIDES, [(None, 100), ("horz", 200)]))

    b = build_brief(open_presentation(make_deck(build)))
    assert b["guides"] == {"vertical": [1.49, 32.37], "horizontal": [4.82, 16.71]}  # the master's, each once
    assert b["content_area_source"] == "guides"


def test_presentation_guides_when_the_master_has_none(make_deck):
    def build(prs):
        prs.slide_width = Cm(33.87)
        _add_ext(prs.part._element, _guide_ext(SLIDE_GUIDES, VERTICAL + HORIZONTAL))
        _add_ext(prs.part._element, _guide_ext(NOTES_GUIDES, [(None, 500)]))

    b = build_brief(open_presentation(make_deck(build)))
    assert b["guides"] == {"vertical": [1.49, 32.37], "horizontal": [4.82, 16.71]}
    assert b["content_area"] == pytest.approx([1.5, 4.82, 30.87, 11.88], abs=0.02)
