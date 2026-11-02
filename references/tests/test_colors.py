import pytest
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn

from pptx_skills.colors import ColorRef, apply_color, format_color, parse_color, read_color
from pptx_skills.errors import PptxSkillsError


@pytest.mark.parametrize("text,ref", [
    ("accent1", ColorRef("accent1", None, 0)), ("accent1+60", ColorRef("accent1", None, 60)),
    ("dk2-25", ColorRef("dk2", None, -25)), ("tx1", ColorRef("tx1", None, 0))])
def test_parse(text, ref):
    assert parse_color(text) == ref and format_color(ref) == text


def test_none_and_errors():
    assert parse_color("none") is None
    with pytest.raises(PptxSkillsError, match="unknown color"):
        parse_color("accent9")
    assert parse_color("#ff0000") == ColorRef(None, "FF0000", 0)
    assert format_color(ColorRef(None, "FF0000", 0)) == "#FF0000"
    with pytest.raises(PptxSkillsError, match="unknown color"):
        parse_color("accent1+100")
    with pytest.raises(PptxSkillsError, match="unknown color"):
        parse_color("#GG0000")


def test_apply_then_read(default_pptx):
    prs = Presentation(default_pptx)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, 100, 100)
    shp.fill.solid()
    apply_color(shp.fill.fore_color, parse_color("accent1+60"))
    assert read_color(shp._element.spPr.find(qn("a:solidFill"))) == "accent1+60"
    apply_color(shp.fill.fore_color, parse_color("dk2-25"))
    assert read_color(shp._element.spPr.find(qn("a:solidFill"))) == "dk2-25"
    apply_color(shp.fill.fore_color, parse_color("#1f4e79"))
    assert read_color(shp._element.spPr.find(qn("a:solidFill"))) == "#1F4E79"


def test_read_lummod_only_and_sysclr():
    el = parse_xml(f'<a:solidFill {nsdecls("a")}><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>')
    assert read_color(el) == "accent1-25"
    el = parse_xml(f'<a:solidFill {nsdecls("a")}><a:sysClr val="windowText" lastClr="000000"/></a:solidFill>')
    assert read_color(el) == "#000000"


def test_read_pattern_fill_and_empty():
    el = parse_xml(f'<a:pattFill {nsdecls("a")} prst="pct5"><a:fgClr><a:schemeClr val="accent2"/></a:fgClr>'
                   '<a:bgClr><a:srgbClr val="ffffff"/></a:bgClr></a:pattFill>')
    assert read_color(el.find(qn("a:fgClr"))) == "accent2"
    assert read_color(el.find(qn("a:bgClr"))) == "#FFFFFF"
    assert read_color(el) is None
    assert read_color(parse_xml(f'<a:solidFill {nsdecls("a")}/>')) is None


def _fill(inner):
    return parse_xml(f'<a:solidFill {nsdecls("a")}>{inner}</a:solidFill>')


@pytest.mark.parametrize("inner,expected", [
    ('<a:sysClr val="windowText"/>', "#000000"),
    ('<a:sysClr val="window"/>', "#FFFFFF"),
    ('<a:sysClr val="btnFace"/>', None),
    ('<a:sysClr val="window" lastClr="f0f0f0"/>', "#F0F0F0"),
    ("<a:srgbClr/>", None),
    ("<a:schemeClr/>", None),
    ('<a:schemeClr val="accent1"><a:lumOff/></a:schemeClr>', "accent1"),
])
def test_read_valid_colors_without_optional_attributes(inner, expected):
    assert read_color(_fill(inner)) == expected


@pytest.mark.parametrize("inner,expected", [
    ('<a:schemeClr val="accent1"><a:lumMod val="99600"/><a:lumOff val="400"/></a:schemeClr>', "accent1"),
    ('<a:schemeClr val="accent1"><a:lumMod val="99600"/></a:schemeClr>', "accent1"),
    ('<a:schemeClr val="accent1"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr>', "accent1+60"),
])
def test_read_never_writes_a_zero_brightness(inner, expected):
    assert read_color(_fill(inner)) == expected


def test_apply_color_replaces_modifiers_but_keeps_transparency():
    el = parse_xml(
        f'<a:solidFill {nsdecls("a")}><a:schemeClr val="accent1"><a:tint val="66000"/><a:satMod val="160000"/>'
        '<a:alpha val="50000"/></a:schemeClr></a:solidFill>'
    )
    from pptx.dml.color import ColorFormat

    apply_color(ColorFormat.from_colorchoice_parent(el), parse_color("accent2+40"))
    assert read_color(el) == "accent2+40"
    assert [c.tag.split("}")[1] for c in el.find(qn("a:schemeClr"))] == ["alpha", "lumMod", "lumOff"]
    apply_color(ColorFormat.from_colorchoice_parent(el), parse_color("#112233"))
    assert [c.tag.split("}")[1] for c in el.find(qn("a:srgbClr"))] == ["alpha"]
