import pytest
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.ns import qn
from pptx.util import Cm

from conftest import apply, layout_named
from pptx_skills.errors import PptxSkillsError
from pptx_skills.inventory import inventory
from pptx_skills.opc import open_presentation


@pytest.fixture
def blank(make_deck):
    return make_deck(lambda prs: prs.slides.add_slide(layout_named(prs, "Blank")))


def shape_named(deck, name, index=0):
    return next(s for s in open_presentation(deck).slides[index].shapes if s.name == name)


@pytest.fixture
def png_200x100(tmp_path):
    from PIL import Image

    path = tmp_path / "wide.png"
    Image.new("RGB", (200, 100), (200, 30, 30)).save(path)
    return path


@pytest.fixture
def svg_file(tmp_path):
    path = tmp_path / "logo.svg"
    path.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>')
    return path


def chart_op(**extra):
    return {"op": "add_chart", "slide": 1, "name": "chart", "box": [2, 3, 20, 10], "type": "column_clustered",
            "categories": ["A", "B"], "series": [{"name": "S", "values": [1, 2]}], **extra}


def test_column_chart_editable(blank):
    apply(blank, [chart_op(series=[{"name": "S1", "values": [1, 2]}, {"name": "S2", "values": [3, 4]}])])
    frame = shape_named(blank, "chart")
    chart = frame.chart
    assert chart.chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED
    plot = chart.plots[0]
    assert list(plot.categories) == ["A", "B"]
    assert [tuple(s.values) for s in plot.series] == [(1.0, 2.0), (3.0, 4.0)]
    fills = [s.format.fill.fore_color.theme_color for s in plot.series]
    assert fills == [MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_2]
    assert chart._chartSpace.xpath("//c:ser//a:srgbClr") == []
    item = next(s for s in inventory(open_presentation(blank))["slides"][0]["shapes"] if s["name"] == "chart")
    assert item["kind"] == "chart"


def test_chart_text_and_decoration(blank):
    apply(blank, [chart_op(labels=True, legend="right")])
    chart = shape_named(blank, "chart").chart
    xml = chart._chartSpace
    assert chart.font.size.pt >= 11
    assert chart.font.color.theme_color == MSO_THEME_COLOR.TEXT_1
    assert chart.plots[0].data_labels.font.size.pt >= 11
    assert chart.legend.font.size.pt >= 11
    assert xml.xpath("//c:view3D | //a:gradFill | //a:effectLst/*") == []
    assert chart.value_axis.has_major_gridlines and chart.value_axis.major_gridlines.format.line.width.pt <= 0.75


def test_chart_types_legend_defaults(blank):
    apply(blank, [
        chart_op(name="line", type="line"),
        chart_op(name="stack", type="bar_stacked", legend="none"),
        chart_op(name="ring", type="doughnut", categories=["A", "B", "C"], series=[{"name": "S", "values": [1, 2, 3]}]),
    ])
    line, stack, ring = (shape_named(blank, n).chart for n in ("line", "stack", "ring"))
    assert line.chart_type == XL_CHART_TYPE.LINE and not line.has_legend
    assert stack.chart_type == XL_CHART_TYPE.BAR_STACKED and not stack.has_legend
    assert ring.chart_type == XL_CHART_TYPE.DOUGHNUT and ring.has_legend
    colors = [ring.plots[0].series[0].points[i].format.fill.fore_color.theme_color for i in range(3)]
    assert colors == [MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_2, MSO_THEME_COLOR.ACCENT_3]
    assert line.plots[0].series[0].format.line.color.theme_color == MSO_THEME_COLOR.ACCENT_1


@pytest.mark.parametrize("bad,message", [
    ({"series": [{"name": "S", "values": [1]}]}, "values"),
    ({"series": [{"name": str(i), "values": [1, 2]} for i in range(7)]}, "accent"),
    ({"name": "dup"}, "already exists"),
])
def test_chart_errors(blank, bad, message):
    apply(blank, [{"op": "add_shape", "slide": 1, "name": "dup", "shape": "rect", "box": [1, 1, 2, 2]}])
    with pytest.raises(PptxSkillsError, match=message):
        apply(blank, [chart_op(**bad)])


def test_picture_cropped_to_fill(blank, png_200x100):
    apply(blank, [{"op": "add_picture", "slide": 1, "name": "pic", "path": str(png_200x100), "box": [2, 2, 4, 4], "alt": "A red field"}])
    pic = shape_named(blank, "pic")
    assert pic.crop_left + pic.crop_right == pytest.approx(0.5)
    assert pic.crop_top == 0 and pic.crop_bottom == 0
    assert (pic.width, pic.height) == (Cm(4), Cm(4))
    assert pic._element.nvPicPr.cNvPr.get("descr") == "A red field"


def test_picture_relative_path_and_tall_crop(blank, png_200x100, monkeypatch):
    monkeypatch.chdir(png_200x100.parent)
    apply(blank, [{"op": "add_picture", "slide": 1, "name": "pic", "path": "wide.png", "box": [2, 2, 8, 2], "alt": "x"}])
    pic = shape_named(blank, "pic")
    assert pic.crop_left == 0 and pic.crop_top == pytest.approx(0.25) and pic.crop_bottom == pytest.approx(0.25)


def test_picture_errors(blank, svg_file, png_200x100, tmp_path):
    base = {"op": "add_picture", "slide": 1, "name": "pic", "box": [2, 2, 4, 4], "alt": "x"}
    with pytest.raises(PptxSkillsError, match="PNG or JPG"):
        apply(blank, [{**base, "path": str(svg_file)}])
    fake = tmp_path / "fake.png"
    fake.write_text("not an image")
    with pytest.raises(PptxSkillsError, match="PNG or JPG"):
        apply(blank, [{**base, "path": str(fake)}])
    with pytest.raises(PptxSkillsError, match="not found"):
        apply(blank, [{**base, "path": str(tmp_path / "missing.png")}])
    with pytest.raises(PptxSkillsError, match="alt"):
        apply(blank, [{k: v for k, v in {**base, "path": str(png_200x100)}.items() if k != "alt"}])
    with pytest.raises(PptxSkillsError):
        apply(blank, [{**base, "path": str(png_200x100), "target": {"role": "picture"}}])


def test_picture_into_placeholder(make_deck, png_200x100):
    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Title Only"))
        from lxml import etree
        from pptx.oxml.ns import nsdecls

        sp = etree.fromstring(
            f'<p:sp {nsdecls("p", "a")}><p:nvSpPr><p:cNvPr id="30" name="Picture Placeholder 1"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="pic" idx="14"/></p:nvPr></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{Cm(2)}" y="{Cm(6)}"/><a:ext cx="{Cm(6)}" cy="{Cm(6)}"/></a:xfrm></p:spPr></p:sp>'
        )
        slide.shapes._spTree.append(sp)

    deck = make_deck(build)
    apply(deck, [{"op": "add_picture", "slide": 1, "name": "photo", "path": str(png_200x100), "target": {"name": "Picture Placeholder 1"}, "alt": "Team"}])
    pic = shape_named(deck, "photo")
    assert pic._element.tag == qn("p:pic") and pic.is_placeholder
    assert pic.crop_left + pic.crop_right == pytest.approx(0.5)
    assert pic._element.nvPicPr.cNvPr.get("descr") == "Team"


def test_picture_target_not_a_picture_placeholder(blank, png_200x100):
    apply(blank, [{"op": "add_shape", "slide": 1, "name": "r", "shape": "rect", "box": [1, 1, 2, 2]}])
    with pytest.raises(PptxSkillsError, match="picture placeholder"):
        apply(blank, [{"op": "add_picture", "slide": 1, "name": "p", "path": str(png_200x100), "target": {"name": "r"}, "alt": "x"}])


def test_notes(blank):
    apply(blank, [{"op": "set_notes", "slide": 1, "text": "Say this"}])
    assert open_presentation(blank).slides[0].notes_slide.notes_text_frame.text == "Say this"
    apply(blank, [{"op": "set_notes", "slide": 1, "text": "One\nTwo"}])
    frame = open_presentation(blank).slides[0].notes_slide.notes_text_frame
    assert [p.text for p in frame.paragraphs] == ["One", "Two"]


@pytest.fixture
def no_notes_body(make_deck):
    """A slide whose notes page has no body placeholder, because the notes master has none."""

    def build(prs):
        master = prs.notes_master
        for ph in list(master.placeholders):
            if ph.placeholder_format.type.name == "BODY":
                ph._element.getparent().remove(ph._element)
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        assert slide.notes_slide.notes_text_frame is None

    return make_deck(build, name="no-notes-body.pptx")


def test_notes_page_without_body_placeholder(no_notes_body, capsys):
    from pptx_skills.checks import run_checks
    from pptx_skills.cli import main

    assert inventory(open_presentation(no_notes_body))["slides"][0]["notes"] == ""
    assert run_checks(no_notes_body, strict=True)["summary"]["error"] == 0
    with pytest.raises(PptxSkillsError, match="op 1 .set_notes.: .*no body placeholder"):
        apply(no_notes_body, [{"op": "set_notes", "slide": 1, "text": "Say this"}])
    assert main(["inventory", str(no_notes_body)]) == 0
