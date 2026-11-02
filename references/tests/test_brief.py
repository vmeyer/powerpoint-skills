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
