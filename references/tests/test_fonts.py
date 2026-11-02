import math
from pathlib import Path

import pytest
from lxml import etree
from PIL import ImageFont
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

from conftest import layout_named
from pptx_skills import fonts
from pptx_skills.fonts import RunStyle, effective_run_style, find_font, measure_paragraphs
from pptx_skills.units import EMU_PER_CM, EMU_PER_PT


def style_of(shape, slide, paragraph=0, run=0):
    p = shape.text_frame.paragraphs[paragraph]
    return effective_run_style(p.runs[run]._r, p._p, shape, slide)


def list_style(shape_or_element):
    element = getattr(shape_or_element, "_element", shape_or_element)
    return element.find(qn("p:txBody")).find(qn("a:lstStyle"))


def tx_style(prs, tag):
    return prs.slide_master._element.find(qn("p:txStyles")).find(qn(tag))


def level_rpr(styles, level=1):
    """The `a:lvlNpPr/a:defRPr` of a list style, created when missing."""
    lvl = styles.find(qn(f"a:lvl{level}pPr"))
    if lvl is None:
        lvl = etree.SubElement(styles, qn(f"a:lvl{level}pPr"))
    def_rpr = lvl.find(qn("a:defRPr"))
    if def_rpr is None:
        def_rpr = etree.SubElement(lvl, qn("a:defRPr"))
    return def_rpr


def master_placeholder(prs, ph_type):
    return next(ph for ph in prs.slide_master.placeholders if ph.placeholder_format.type == ph_type)


def set_theme_fonts(prs, major, minor):
    part = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(part.blob)
    for tag, typeface in (("a:majorFont", major), ("a:minorFont", minor)):
        root.find(f".//{qn(tag)}/{qn('a:latin')}").set("typeface", typeface)
    part._blob = etree.tostring(root)


def to_cm(pt):
    return pt * EMU_PER_PT / EMU_PER_CM


@pytest.fixture
def aileron(tmp_path, monkeypatch):
    """A font folder holding Pillow's bundled Aileron Regular; the lookup searches only this folder."""
    font_dir = tmp_path / "fonts"
    font_dir.mkdir()
    path = font_dir / "Aileron-Regular.ttf"
    path.write_bytes(ImageFont.load_default(10).font_bytes)
    monkeypatch.setattr(fonts, "_FONT_DIRS", (str(font_dir),))
    return path


# --- effective style -----------------------------------------------------------------------------


def test_effective_style_inherits_from_master(default_pptx):
    prs = Presentation(default_pptx); s = prs.slides.add_slide(prs.slide_layouts[1]); s.placeholders[1].text = "x"
    p = s.placeholders[1].text_frame.paragraphs[0]
    st = effective_run_style(p.runs[0]._r, p._p, s.placeholders[1], s)
    assert (st.family, st.size_pt) == ("Calibri", 32)


def test_title_uses_title_style_and_major_font(default_pptx):
    prs = Presentation(default_pptx)
    set_theme_fonts(prs, "Major Sans", "Minor Serif")
    s = prs.slides.add_slide(layout_named(prs, "Title and Content"))
    s.shapes.title.text, s.placeholders[1].text = "T", "B"
    assert style_of(s.shapes.title, s) == RunStyle("Major Sans", 44, False, False)
    assert style_of(s.placeholders[1], s) == RunStyle("Minor Serif", 32, False, False)
    cover = prs.slides.add_slide(layout_named(prs, "Title Slide"))  # ctrTitle and subTitle
    cover.shapes.title.text, cover.placeholders[1].text = "T", "S"
    assert style_of(cover.shapes.title, cover).size_pt == 44
    assert style_of(cover.placeholders[1], cover).size_pt == 32


def test_placeholder_chain_first_explicit_value_wins(default_pptx):
    prs = Presentation(default_pptx)
    layout = layout_named(prs, "Title and Content")
    s = prs.slides.add_slide(layout)
    body = s.placeholders[1]
    body.text = "x"
    assert style_of(body, s).size_pt == 32  # master bodyStyle
    steps = (
        (list_style(master_placeholder(prs, PP_PLACEHOLDER.BODY)), 20),
        (list_style(layout.placeholders.get(idx=1)), 22),
        (list_style(body), 24),
    )
    for styles, size in steps:
        level_rpr(styles).set("sz", str(size * 100))
        assert style_of(body, s).size_pt == size
    p_pr = body.text_frame.paragraphs[0]._p.get_or_add_pPr()
    etree.SubElement(p_pr, qn("a:defRPr")).set("sz", "2600")
    assert style_of(body, s).size_pt == 26
    body.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
    assert style_of(body, s).size_pt == 28


def test_textbox_uses_other_style_then_presentation_default_then_18(default_pptx):
    prs = Presentation(default_pptx)
    s = prs.slides.add_slide(layout_named(prs, "Blank"))
    box = s.shapes.add_textbox(Cm(1), Cm(1), Cm(5), Cm(2))
    box.text_frame.text = "x"
    other = level_rpr(tx_style(prs, "p:otherStyle"))
    default = level_rpr(prs._element.find(qn("p:defaultTextStyle")))
    other.set("sz", "1500")
    assert style_of(box, s).size_pt == 15
    del other.attrib["sz"]
    default.set("sz", "1600")
    assert style_of(box, s).size_pt == 16
    del default.attrib["sz"]
    assert style_of(box, s) == RunStyle("Calibri", 18, False, False)


def test_paragraph_level_selects_list_style_level(default_pptx):
    prs = Presentation(default_pptx)
    s = prs.slides.add_slide(layout_named(prs, "Title and Content"))
    tf = s.placeholders[1].text_frame
    tf.text = "zero\none\ntwo"
    for level, p in enumerate(tf.paragraphs):
        p.level = level
    assert [style_of(s.placeholders[1], s, paragraph=i).size_pt for i in range(3)] == [32, 28, 24]


def test_theme_tokens_and_title_default_to_major(default_pptx):
    prs = Presentation(default_pptx)
    set_theme_fonts(prs, "Major Sans", "Minor Serif")
    for styles in (prs.slide_master._element.find(qn("p:txStyles")), prs._element.find(qn("p:defaultTextStyle"))):
        for latin in list(styles.iter(qn("a:latin"))):
            latin.getparent().remove(latin)
    s = prs.slides.add_slide(layout_named(prs, "Title and Content"))
    s.shapes.title.text, s.placeholders[1].text = "T", "B"
    assert style_of(s.shapes.title, s).family == "Major Sans"
    assert style_of(s.placeholders[1], s).family == "Minor Serif"
    run = s.placeholders[1].text_frame.paragraphs[0].runs[0]
    run.font.name = "+mj-lt"
    assert style_of(s.placeholders[1], s).family == "Major Sans"
    run.font.name = "Explicit Face"
    assert style_of(s.placeholders[1], s).family == "Explicit Face"


def test_bold_and_italic_follow_the_chain(default_pptx):
    prs = Presentation(default_pptx)
    level_rpr(tx_style(prs, "p:titleStyle")).set("b", "1")
    s = prs.slides.add_slide(layout_named(prs, "Title Only"))
    s.shapes.title.text = "T"
    assert (style_of(s.shapes.title, s).bold, style_of(s.shapes.title, s).italic) == (True, False)
    run = s.shapes.title.text_frame.paragraphs[0].runs[0]
    run.font.bold, run.font.italic = False, True
    assert (style_of(s.shapes.title, s).bold, style_of(s.shapes.title, s).italic) == (False, True)


def test_layout_placeholder_maps_to_master_by_type_not_idx(default_pptx):
    prs = Presentation(default_pptx)
    layout = layout_named(prs, "Two Content")
    right = list_style(layout.placeholders.get(idx=2))
    for child in list(right):
        right.remove(child)
    s = prs.slides.add_slide(layout)
    s.placeholders[2].text = "x"
    # the master's date placeholder also has idx 2 (12 pt); an object placeholder inherits from the master body
    assert style_of(s.placeholders[2], s).size_pt == 32


def test_slide_placeholder_falls_back_to_layout_type_when_idx_is_unknown(default_pptx):
    prs = Presentation(default_pptx)
    layout = layout_named(prs, "Title and Content")
    level_rpr(list_style(layout.placeholders.get(idx=1))).set("sz", "2200")
    s = prs.slides.add_slide(layout)
    body = s.placeholders[1]
    body.text = "x"
    body._element.ph.set("idx", "7")
    assert style_of(body, s).size_pt == 22


# --- font lookup ---------------------------------------------------------------------------------


def test_find_font_matches_family_name_and_caches(aileron):
    assert find_font("Aileron") == aileron
    assert find_font("aileron", bold=True, italic=True) == aileron  # no bold italic face: the regular one
    assert find_font("Aileron Regular") == aileron  # full name, as PowerPoint names non-RIBBI faces
    assert find_font("NoSuchFont") is None
    aileron.unlink()
    assert find_font("Aileron") == aileron  # the folders are scanned once per process


def test_find_font_prefers_the_requested_style(monkeypatch):
    def install(*styles):
        faces = [fonts._Face(Path(f"/fonts/Sans-{style.replace(' ', '')}.ttf"), 0, style) for style in styles]
        monkeypatch.setattr(fonts, "_font_index", lambda dirs: ({"sans": faces}, {}))

    install("Light", "Semibold", "Regular", "Light Italic", "Bold Italic", "Bold", "Italic")
    assert find_font("Sans").name == "Sans-Regular.ttf"
    assert find_font("Sans", bold=True).name == "Sans-Bold.ttf"
    assert find_font("Sans", italic=True).name == "Sans-Italic.ttf"
    assert find_font("Sans", bold=True, italic=True).name == "Sans-BoldItalic.ttf"
    install("Light", "Regular", "Light Italic")
    assert find_font("Sans", bold=True).name == "Sans-Regular.ttf"  # no bold face: the plain one
    assert find_font("Sans", italic=True).name == "Sans-LightItalic.ttf"


# --- measurement ---------------------------------------------------------------------------------


def test_wrap_counts_lines_with_default_font(slide_with_box):
    slide, textbox = slide_with_box
    box = textbox(width_cm=5, text="word " * 40, size=12)
    m = measure_paragraphs(box, slide)
    assert m.lines_per_paragraph[0] > 3 and m.height_cm > 2


def test_hard_break_and_missing_font(slide_with_box):
    slide, textbox = slide_with_box
    box = textbox(width_cm=20, text="a\vb", size=12, family="NoSuchFont")
    m = measure_paragraphs(box, slide)
    assert m.lines_per_paragraph == [2] and "NoSuchFont" in m.fonts_missing


def test_vertical_tab_in_run_text_is_a_line_break(slide_with_box):
    slide, textbox = slide_with_box
    box = textbox(width_cm=20, text="x", size=12)
    box.text_frame.paragraphs[0].runs[0].text = "a\vb"  # python-pptx stores it as _x000B_
    assert measure_paragraphs(box, slide).lines_per_paragraph == [2]


def test_installed_font_is_not_missing(slide_with_box, aileron):
    slide, textbox = slide_with_box
    box = textbox(width_cm=20, text="a b", size=12, family="Aileron")
    assert measure_paragraphs(box, slide).fonts_missing == set()


def test_lines_are_counted_per_paragraph(slide_with_box):
    slide, textbox = slide_with_box
    box = textbox(width_cm=20, text="one\n\nthree", size=12)
    assert measure_paragraphs(box, slide).lines_per_paragraph == [1, 1, 1]


def test_no_wrap_counts_hard_breaks_only(slide_with_box):
    slide, textbox = slide_with_box
    box = textbox(width_cm=3, text="word " * 40 + "\vend", size=12)
    box.text_frame.word_wrap = False
    assert measure_paragraphs(box, slide).lines_per_paragraph == [2]


def test_line_height_spacing_and_insets(slide_with_box):
    slide, textbox = slide_with_box
    box = textbox(width_cm=20, text="Line", size=12)
    tf = box.text_frame
    assert measure_paragraphs(box, slide).height_cm == round(to_cm(12 * 1.2) + 2 * 45720 / EMU_PER_CM, 2)  # 0.76
    tf.margin_top = tf.margin_bottom = 0
    assert measure_paragraphs(box, slide).height_cm == round(to_cm(14.4), 2)
    p = tf.paragraphs[0]
    p.line_spacing = 2.0
    assert measure_paragraphs(box, slide).height_cm == round(to_cm(28.8), 2)
    p.line_spacing = Pt(30)
    assert measure_paragraphs(box, slide).height_cm == round(to_cm(30), 2)
    p.space_before, p.space_after = Pt(6), Pt(4)
    assert measure_paragraphs(box, slide).height_cm == round(to_cm(40), 2)
    tf.margin_top = Cm(1)
    assert measure_paragraphs(box, slide).height_cm == round(1 + to_cm(40), 2)


def test_paragraph_spacing_inherits_from_master_body_style(default_pptx):
    prs = Presentation(default_pptx)
    s = prs.slides.add_slide(layout_named(prs, "Title and Content"))
    body = s.placeholders[1]
    body.text = "x"
    line = 32 * 1.2
    expected = to_cm(line + 0.2 * line) + 2 * 45720 / EMU_PER_CM  # bodyStyle spcBef 20 % of a line, master insets
    assert measure_paragraphs(body, s).height_cm == round(expected, 2)


def test_placeholder_insets_inherit_from_master(default_pptx):
    prs = Presentation(default_pptx)
    s = prs.slides.add_slide(layout_named(prs, "Title and Content"))
    body = s.placeholders[1]
    body.text = "x"
    before = measure_paragraphs(body, s).height_cm
    master_body_pr = master_placeholder(prs, PP_PLACEHOLDER.BODY)._element.find(qn("p:txBody")).find(qn("a:bodyPr"))
    master_body_pr.set("tIns", "0")
    master_body_pr.set("bIns", "0")
    after = measure_paragraphs(body, s).height_cm
    assert before - after == pytest.approx(2 * 45720 / EMU_PER_CM, abs=0.011)
    body.text_frame.margin_top = Cm(1)  # the slide's own bodyPr wins
    assert measure_paragraphs(body, s).height_cm == pytest.approx(after + 1, abs=0.011)


def test_widest_word_and_overwide_word(slide_with_box):
    slide, textbox = slide_with_box
    word = "w" * 40
    box = textbox(width_cm=3, text="a " + word, size=12)
    m = measure_paragraphs(box, slide)
    word_pt = ImageFont.load_default(120).getlength(word) / 10
    available_pt = (3 * EMU_PER_CM - 2 * 91440) / EMU_PER_PT
    assert m.widest_word_cm == round(to_cm(word_pt), 2)
    assert m.lines_per_paragraph == [1 + math.ceil(word_pt / available_pt)]  # "a", then the word broken across lines


def test_left_margin_narrows_the_wrap_width(slide_with_box):
    slide, textbox = slide_with_box
    box = textbox(width_cm=10, text="aa aa aa aa", size=12)
    assert measure_paragraphs(box, slide).lines_per_paragraph == [1]
    p_pr = box.text_frame.paragraphs[0]._p.get_or_add_pPr()
    p_pr.set("marL", str(Cm(10) - 2 * 91440 - Cm(0.6)))  # leaves 0.6 cm: room for one "aa" per line
    assert measure_paragraphs(box, slide).lines_per_paragraph == [4]


def test_shape_without_text_measures_empty(default_pptx):
    from pptx.enum.shapes import MSO_CONNECTOR

    prs = Presentation(default_pptx)
    s = prs.slides.add_slide(layout_named(prs, "Blank"))
    line = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(1), Cm(1), Cm(5), Cm(1))
    m = measure_paragraphs(line, s)
    assert (m.lines_per_paragraph, m.height_cm, m.widest_word_cm, m.fonts_missing) == ([], 0.0, 0.0, set())
