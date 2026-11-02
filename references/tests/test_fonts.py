import math
import struct
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


def body_pr(shape_or_element):
    element = getattr(shape_or_element, "_element", shape_or_element)
    return element.find(qn("p:txBody")).find(qn("a:bodyPr"))


DEFAULT_FONT_DIRS = fonts._FONT_DIRS  # captured before the fixture below hides installed fonts


@pytest.fixture(autouse=True)
def no_installed_fonts(monkeypatch):
    """Lookups see no installed fonts unless a test installs some, so results do not depend on the machine.

    Missing fonts are measured with Pillow's default font and a single line of 1.2 × size.
    """
    monkeypatch.setattr(fonts, "_FONT_DIRS", ())


@pytest.fixture
def aileron(tmp_path, monkeypatch, no_installed_fonts):
    """A font folder holding Pillow's bundled Aileron Regular; the lookup searches only this folder."""
    font_dir = tmp_path / "fonts"
    font_dir.mkdir()
    path = font_dir / "Aileron-Regular.ttf"
    path.write_bytes(ImageFont.load_default(10).font_bytes)
    monkeypatch.setattr(fonts, "_FONT_DIRS", (str(tmp_path / "absent"), str(font_dir)))  # absent folders are skipped
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


def test_font_search_covers_office_folders():
    office = "Microsoft PowerPoint.app/Contents/Resources/DFonts"
    assert f"/Applications/{office}" in DEFAULT_FONT_DIRS
    assert f"~/Applications/{office}" in DEFAULT_FONT_DIRS


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


def set_line_extent(path, ascent_em, descent_em):
    """Rewrite a font's hhea ascender and descender, the line extent FreeType reports."""
    data = bytearray(path.read_bytes())
    count = struct.unpack(">H", data[4:6])[0]
    offsets = {bytes(data[12 + 16 * i : 16 + 16 * i]): struct.unpack(">I", data[20 + 16 * i : 24 + 16 * i])[0] for i in range(count)}
    units = struct.unpack(">H", data[offsets[b"head"] + 18 : offsets[b"head"] + 20])[0]
    struct.pack_into(">hh", data, offsets[b"hhea"] + 4, round(ascent_em * units), -round(descent_em * units))
    path.write_bytes(bytes(data))


def test_line_height_uses_the_font_metrics(slide_with_box, aileron):
    set_line_extent(aileron, 1.5, 0.5)  # a line of 2 × size; Aileron itself is 1.2, like the fallback
    assert sum(ImageFont.truetype(str(aileron), 1000).getmetrics()) == 2000
    slide, textbox = slide_with_box
    installed = textbox(width_cm=20, text="Line", size=100, family="Aileron")
    missing = textbox(width_cm=20, text="Line", size=100, family="NoSuchFont")
    for box in (installed, missing):
        box.text_frame.margin_top = box.text_frame.margin_bottom = 0
    assert measure_paragraphs(installed, slide).height_cm == round(to_cm(200), 2)  # 7.06
    assert measure_paragraphs(missing, slide).height_cm == round(to_cm(100 * 1.2), 2)  # 4.23


def test_line_height_spacing_and_insets(slide_with_box):
    slide, textbox = slide_with_box
    box = textbox(width_cm=20, text="Line", size=12)  # Calibri is not installed: a line is 1.2 × 12 pt
    tf = box.text_frame
    assert measure_paragraphs(box, slide).height_cm == round(to_cm(12 * 1.2) + 2 * 45720 / EMU_PER_CM, 2)  # 0.76
    tf.margin_top = tf.margin_bottom = 0
    assert measure_paragraphs(box, slide).height_cm == round(to_cm(14.4), 2)
    p = tf.paragraphs[0]
    p.line_spacing = 2.0
    assert measure_paragraphs(box, slide).height_cm == round(to_cm(28.8), 2)
    p.line_spacing = Pt(30)
    assert measure_paragraphs(box, slide).height_cm == round(to_cm(30), 2)
    p.space_before, p.space_after = Pt(6), Pt(4)  # outside the only paragraph: not counted
    assert measure_paragraphs(box, slide).height_cm == round(to_cm(30), 2)
    tf.margin_top = Cm(1)
    assert measure_paragraphs(box, slide).height_cm == round(1 + to_cm(30), 2)


def test_spacing_counts_between_paragraphs_only(slide_with_box):
    slide, textbox = slide_with_box
    box = textbox(width_cm=20, text="One\nTwo", size=12)
    box.text_frame.margin_top = box.text_frame.margin_bottom = 0
    first, second = box.text_frame.paragraphs
    first.space_before, first.space_after = Pt(6), Pt(4)
    second.space_before, second.space_after = Pt(5), Pt(8)
    assert measure_paragraphs(box, slide).height_cm == round(to_cm(2 * 14.4 + 4 + 5), 2)


def test_paragraph_spacing_inherits_from_master_body_style(default_pptx):
    prs = Presentation(default_pptx)
    s = prs.slides.add_slide(layout_named(prs, "Title and Content"))
    body = s.placeholders[1]
    body.text = "x\ny"
    line = 32 * 1.2
    expected = to_cm(2 * line + 0.2 * line) + 2 * 45720 / EMU_PER_CM  # bodyStyle spcBef 20 % of a line, master insets
    assert measure_paragraphs(body, s).height_cm == round(expected, 2)


def test_outer_spacing_keeps_a_kicker_inside_its_box(kicker_master):
    """Calibrated on a reference slide: a 14 pt kicker with 6 pt before and after fits a 0.60 cm box."""
    prs = Presentation(kicker_master)
    master_body = master_placeholder(prs, PP_PLACEHOLDER.BODY)
    for name in ("lIns", "tIns", "rIns", "bIns"):
        body_pr(master_body).set(name, "0")
    level = level_rpr(tx_style(prs, "p:bodyStyle")).getparent()
    level.find(qn("a:defRPr")).set("sz", "1400")
    for tag in ("a:spcAft", "a:spcBef"):  # inserted in front, so spcBef ends up first
        for old in level.findall(qn(tag)):
            level.remove(old)
        spacing = etree.Element(qn(tag))
        etree.SubElement(spacing, qn("a:spcPts")).set("val", "600")
        level.insert(0, spacing)
    s = prs.slides.add_slide(layout_named(prs, "Title Only"))
    kicker = next(ph for ph in s.placeholders if ph.placeholder_format.idx == 13)
    assert len(body_pr(kicker)) == 0 and not body_pr(kicker).attrib  # an empty <a:bodyPr/>
    kicker.text = "Kicker"
    m = measure_paragraphs(kicker, s)
    assert m.height_cm == round(to_cm(14 * 1.2), 2) <= 0.60


def test_layout_supplies_body_properties(default_pptx):
    prs = Presentation(default_pptx)
    layout = layout_named(prs, "Title and Content")
    s = prs.slides.add_slide(layout)
    body = s.placeholders[1]
    body.text = "x"
    m = measure_paragraphs(body, s)
    assert (m.height_cm, m.autofit) == (round(to_cm(32 * 1.2) + 2 * 45720 / EMU_PER_CM, 2), "shrink")  # from the master
    layout_body_pr = body_pr(layout.placeholders.get(idx=1))
    layout_body_pr.set("tIns", "0")
    layout_body_pr.set("bIns", "0")
    etree.SubElement(layout_body_pr, qn("a:noAutofit"))
    m = measure_paragraphs(body, s)
    assert (m.height_cm, m.autofit) == (round(to_cm(32 * 1.2), 2), "none")


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


def test_autofit_is_resolved_and_exposed(slide_with_box):
    slide, textbox = slide_with_box
    box = textbox(width_cm=5, text="x", size=12)

    def fit():
        m = measure_paragraphs(box, slide)
        return m.autofit, m.font_scale, m.line_spacing_reduction

    assert fit() == ("shape", 1.0, 0.0)  # python-pptx text boxes carry spAutoFit
    for child in list(body_pr(box)):
        body_pr(box).remove(child)
    assert fit() == ("none", 1.0, 0.0)
    norm = etree.SubElement(body_pr(box), qn("a:normAutofit"))
    assert fit() == ("shrink", 1.0, 0.0)
    norm.set("fontScale", "62500")
    norm.set("lnSpcReduction", "20000")
    assert fit() == ("shrink", 0.625, 0.2)
    norm.set("fontScale", "70%")
    assert fit() == ("shrink", 0.7, 0.2)


def test_unreadable_font_file_is_measured_as_missing(slide_with_box, aileron):
    slide, textbox = slide_with_box
    assert find_font("Aileron") == aileron  # indexed while still readable
    aileron.write_bytes(b"not a font any more")
    box = textbox(width_cm=20, text="a b", size=12, family="Aileron")
    m = measure_paragraphs(box, slide)
    assert (m.fonts_missing, m.lines_per_paragraph) == ({"Aileron"}, [1])
    aileron.unlink()
    assert measure_paragraphs(box, slide).fonts_missing == {"Aileron"}


def test_malformed_numbers_count_as_absent(default_pptx):
    prs = Presentation(default_pptx)
    s = prs.slides.add_slide(layout_named(prs, "Title and Content"))
    body = s.placeholders[1]
    body.text = "x"
    p = body.text_frame.paragraphs[0]
    p.runs[0]._r.get_or_add_rPr().set("sz", "large")
    p._p.get_or_add_pPr().set("lvl", "one")
    p._p.get_or_add_pPr().set("marL", "wide")
    body_pr(body).set("tIns", "thin")
    assert style_of(body, s).size_pt == 32
    m = measure_paragraphs(body, s)
    assert (m.lines_per_paragraph, m.height_cm) == ([1], round(to_cm(32 * 1.2) + 2 * 45720 / EMU_PER_CM, 2))
    body._element.ph.set("idx", "first")
    assert style_of(body, s).size_pt == 32


def test_shape_without_text_measures_empty(default_pptx):
    from pptx.enum.shapes import MSO_CONNECTOR

    prs = Presentation(default_pptx)
    s = prs.slides.add_slide(layout_named(prs, "Blank"))
    line = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(1), Cm(1), Cm(5), Cm(1))
    m = measure_paragraphs(line, s)
    assert (m.lines_per_paragraph, m.height_cm, m.widest_word_cm, m.fonts_missing) == ([], 0.0, 0.0, set())
