import json

from pptx import Presentation
from pptx.enum.dml import MSO_PATTERN_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Cm, Pt

from conftest import layout_named
from pptx_skills.cli import main
from pptx_skills.geometry import Box, absolute_box
from pptx_skills.inventory import inventory, to_markdown
from pptx_skills.roles import placeholder_role
from pptx.enum.shapes import PP_PLACEHOLDER as PH


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def test_title_role_and_ids(default_pptx):
    prs = Presentation(default_pptx)
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Hello world"
    sl = inventory(prs)["slides"][0]
    assert (sl["number"], sl["id"], sl["layout"], sl["title"]) == (1, 256, "Title Only", "Hello world")
    assert sl["kicker"] is None and sl["notes"] == ""
    title = next(x for x in sl["shapes"] if x["role"] == "title")
    assert title["box"] == [1.27, 0.76, 22.86, 3.17]
    assert title["kind"] == "placeholder"


def test_runs_and_fill(default_pptx):
    prs = Presentation(default_pptx)
    s = _blank(prs)
    shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1), Cm(1), Cm(4), Cm(2))
    shp.fill.solid()
    shp.fill.fore_color.theme_color = 5  # ACCENT_1
    shp.fill.fore_color.brightness = 0.6
    p = shp.text_frame.paragraphs[0]
    r1 = p.add_run()
    r1.text = "Bold"
    r1.font.bold = True
    r2 = p.add_run()
    r2.text = " rest"
    shape = inventory(prs)["slides"][0]["shapes"][0]
    assert shape["fill"] == "accent1+60"
    assert shape["geometry"] == "rect" and shape["kind"] == "autoshape"
    assert shape["text"]["paragraphs"][0]["runs"] == [{"t": "Bold", "b": True}, {"t": " rest"}]


def test_run_attributes_line_and_autofit(default_pptx):
    prs = Presentation(default_pptx)
    s = _blank(prs)
    tb = s.shapes.add_textbox(Cm(1), Cm(1), Cm(4), Cm(2))
    shrink = s.shapes.add_textbox(Cm(8), Cm(1), Cm(4), Cm(2))
    shrink.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = "x"
    run.font.italic = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    rect = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1), Cm(5), Cm(2), Cm(2))
    rect.line.color.rgb = RGBColor(255, 0, 0)
    rect.line.width = Pt(2)
    none = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(5), Cm(5), Cm(2), Cm(2))
    none.line.fill.background()
    none.fill.background()
    shapes = inventory(prs)["slides"][0]["shapes"]
    assert shapes[0]["kind"] == "textbox"
    assert shapes[0]["text"]["paragraphs"][0]["runs"] == [{"t": "x", "i": True, "color": "#1F4E79", "size": 14}]
    assert shapes[0]["text"]["autofit"] == "shape"
    assert shapes[1]["text"]["autofit"] == "shrink"
    assert shapes[0]["line"] is None and shapes[0]["fill"] == "none"
    assert shapes[2]["line"] == {"color": "#FF0000", "width_pt": 2}
    assert shapes[3]["line"] == "none" and shapes[3]["fill"] == "none"


def test_pattern_fill(default_pptx):
    prs = Presentation(default_pptx)
    shp = _blank(prs).shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Cm(2), Cm(2))
    shp.fill.patterned()
    shp.fill.pattern = MSO_PATTERN_TYPE.WIDE_UPWARD_DIAGONAL
    shp.fill.fore_color.theme_color = 5
    shp.fill.back_color.theme_color = 14  # BACKGROUND_1
    assert inventory(prs)["slides"][0]["shapes"][0]["fill"] == "pattern:wdUpDiag accent1 bg1"


def test_group_children_absolute_boxes(default_pptx):
    prs = Presentation(default_pptx)
    s = _blank(prs)
    grp = s.shapes.add_group_shape()
    grp.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Cm(2), Cm(1))
    grp.left, grp.top, grp.width, grp.height = Cm(5), Cm(5), Cm(4), Cm(2)
    sl = inventory(prs)["slides"][0]
    group = sl["shapes"][0]
    assert group["kind"] == "group" and group["box"] == [5.0, 5.0, 4.0, 2.0]
    assert group["children"][0]["box"] == [5.0, 5.0, 4.0, 2.0]


def test_nested_group_absolute_boxes(default_pptx):
    prs = Presentation(default_pptx)
    s = _blank(prs)
    outer = s.shapes.add_group_shape()
    inner = outer.shapes.add_group_shape()
    child = inner.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1), Cm(1), Cm(2), Cm(2))
    inner.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(3), Cm(3), Cm(2), Cm(2))  # inner child space 1..5
    # inner: child space (1,1,4,4) -> placed in outer child space at (10,10) scaled 2x
    inner.left, inner.top, inner.width, inner.height = Cm(10), Cm(10), Cm(8), Cm(8)
    outer.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Cm(1), Cm(1))  # extends outer child extent
    # outer child space: (0,0)-(18,18); map to slide at (2,2) with scale 0.5
    outer.left, outer.top, outer.width, outer.height = Cm(2), Cm(2), Cm(9), Cm(9)
    # child (1,1,2,2): inner x 10+(1-1)*2=10, w 4 ; outer x 2+10*.5=7, w 2
    assert absolute_box(child) == Box(7.0, 7.0, 2.0, 2.0)
    kids = inventory(prs)["slides"][0]["shapes"][0]["children"]
    inner_inv = next(k for k in kids if k["kind"] == "group")
    assert inner_inv["children"][0]["box"] == [7.0, 7.0, 2.0, 2.0]


def test_kicker_role(kicker_master):
    prs = Presentation(kicker_master)
    prs.slides.add_slide(layout_named(prs, "Title Only"))
    sl = inventory(prs)["slides"][0]
    kicker = next(x for x in sl["shapes"] if x["role"] == "kicker")
    assert kicker["box"] == [1.27, 0.5, 10.0, 0.6]
    assert sl["kicker"] == ""


def test_placeholder_role_rules():
    title = Box(1, 3, 20, 3)
    assert placeholder_role(PH.TITLE, title, title) == "title"
    assert placeholder_role(PH.CENTER_TITLE, title, None) == "title"
    assert placeholder_role(PH.SUBTITLE, title, title) == "subtitle"
    assert placeholder_role(PH.BODY, Box(1, 1, 10, 1.5), title) == "kicker"
    assert placeholder_role(PH.OBJECT, Box(1, 1, 10, 1.6), title) == "body"
    assert placeholder_role(PH.BODY, Box(1, 2.0, 10, 1.0), title) == "kicker"
    assert placeholder_role(PH.BODY, Box(1, 2.5, 10, 1.0), title) == "body"
    assert placeholder_role(PH.BODY, Box(1, 1, 10, 1), None) == "body"
    assert placeholder_role(PH.PICTURE, title, title) == "picture"
    assert placeholder_role(PH.TABLE, title, title) == "table"
    assert placeholder_role(PH.CHART, title, title) == "chart"
    assert placeholder_role(PH.FOOTER, title, title) == "footer"
    assert placeholder_role(PH.SLIDE_NUMBER, title, title) == "number"
    assert placeholder_role(PH.DATE, title, title) == "date"
    assert placeholder_role(PH.MEDIA_CLIP, title, title) == "other"


def test_notes_and_kind_variants(default_pptx):
    prs = Presentation(default_pptx)
    s = _blank(prs)
    s.notes_slide.notes_text_frame.text = "Remember this"
    s.shapes.add_table(2, 2, Cm(1), Cm(1), Cm(6), Cm(2))
    s.shapes.add_connector(1, Cm(1), Cm(5), Cm(4), Cm(6))
    kinds = [x["kind"] for x in inventory(prs)["slides"][0]["shapes"]]
    assert kinds == ["table", "connector"]
    assert inventory(prs)["slides"][0]["notes"] == "Remember this"


def test_filter_markdown_and_cli(deck_with_two_slides, capsys):
    assert [s["number"] for s in inventory(Presentation(deck_with_two_slides), slides=[2])["slides"]] == [2]
    assert main(["inventory", str(deck_with_two_slides), "--json"]) == 0
    assert json.loads(capsys.readouterr().out)["slides"][1]["number"] == 2
    md = to_markdown(inventory(Presentation(deck_with_two_slides)))
    assert "## Slide 1 (id 256) — Title Only" in md
    assert '/title @ 1.27,0.76 22.86×3.17: "First slide"' in md
    assert main(["inventory", str(deck_with_two_slides), "--slides", "2"]) == 0
    out = capsys.readouterr().out
    assert "## Slide 2" in out and "## Slide 1" not in out


def test_cli_errors(deck_with_two_slides, capsys):
    assert main(["inventory", str(deck_with_two_slides), "--slides", "9"]) == 2
    assert "error:" in capsys.readouterr().err
    assert main(["inventory", str(deck_with_two_slides.with_name("missing.pptx"))]) == 2


def test_line_break_and_field_runs(default_pptx):
    from lxml import etree
    from pptx.oxml.ns import nsdecls

    prs = Presentation(default_pptx)
    s = prs.slides.add_slide(prs.slide_layouts[5])
    p = s.shapes.title.text_frame.paragraphs[0]._p
    for xml in (
        '<a:r %s><a:t>Foo</a:t></a:r>',
        '<a:br %s/>',
        '<a:r %s><a:t>Bar</a:t></a:r>',
        '<a:fld %s id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:t>7</a:t></a:fld>',
    ):
        p.append(etree.fromstring(xml % nsdecls("a")))
    sl = inventory(prs)["slides"][0]
    assert sl["title"] == "Foo\nBar7"
    runs = sl["shapes"][0]["text"]["paragraphs"][0]["runs"]
    assert runs == [
        {"t": "Foo"},
        {"t": "\n", "br": True},
        {"t": "Bar"},
        {"t": "7", "field": "slidenum"},
    ]
    assert '"Foo / Bar7"' in to_markdown(inventory(prs))


def test_repeated_main_registers_commands_once(deck_with_two_slides):
    from pptx_skills.cli import COMMAND_MODULES

    for _ in range(2):
        assert main(["inventory", "--help"]) == 0
    assert len(COMMAND_MODULES) == len(set(COMMAND_MODULES))
