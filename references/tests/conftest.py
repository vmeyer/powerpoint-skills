import io
import zipfile

import pytest
from pptx import Presentation

_MAIN = b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
_TEMPLATE = b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"


@pytest.fixture
def default_pptx(tmp_path):
    path = tmp_path / "default.pptx"
    Presentation().save(str(path))
    return path


@pytest.fixture
def potx_file(tmp_path, default_pptx):
    path = tmp_path / "template.potx"
    with zipfile.ZipFile(default_pptx) as src, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(_MAIN, _TEMPLATE)
            dst.writestr(item, data)
    return path


@pytest.fixture
def make_deck(tmp_path):
    """Factory: make_deck(builder, name="deck.pptx") builds on the default template and returns the path."""

    def factory(builder=None, name="deck.pptx"):
        prs = Presentation()
        if builder:
            builder(prs)
        path = tmp_path / name
        prs.save(str(path))
        return path

    return factory


def layout_named(prs, name):
    return next(layout for layout in prs.slide_layouts if layout.name == name)


@pytest.fixture
def deck_with_two_slides(make_deck):
    def build(prs):
        for text in ("First slide", "Second slide"):
            slide = prs.slides.add_slide(layout_named(prs, "Title Only"))
            slide.shapes.title.text = text

    return make_deck(build, name="two.pptx")


@pytest.fixture
def kicker_master(make_deck):
    """Default template whose "Title Only" layout has a 0.6 cm body placeholder above a lowered title."""
    from lxml import etree
    from pptx.oxml.ns import nsdecls
    from pptx.util import Cm

    def build(prs):
        layout = layout_named(prs, "Title Only")
        title = layout.placeholders[0]
        title.left, title.top, title.width, title.height = Cm(1.27), Cm(1.5), Cm(22.86), Cm(3.17)
        sp = etree.fromstring(
            f'<p:sp {nsdecls("p", "a")}><p:nvSpPr><p:cNvPr id="20" name="Kicker"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="body" sz="quarter" idx="13"/></p:nvPr></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{Cm(1.27)}" y="{Cm(0.5)}"/><a:ext cx="{Cm(10)}" cy="{Cm(0.6)}"/></a:xfrm></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Kicker</a:t></a:r></a:p></p:txBody></p:sp>'
        )
        layout.shapes._spTree.append(sp)

    return make_deck(build, name="kicker.pptx")


@pytest.fixture
def guided_master(make_deck):
    """Widescreen default template whose master carries p15 guides: vertical 339/7340, horizontal 1094/3789."""
    from lxml import etree
    from pptx.util import Cm

    ns = "http://schemas.microsoft.com/office/powerpoint/2012/main"

    def build(prs):
        prs.slide_width = Cm(33.87)
        guides = "".join(f'<p15:guide id="{i}" pos="{p}"/>' for i, p in enumerate((339, 7340), 1))
        guides += "".join(f'<p15:guide id="{i}" orient="horz" pos="{p}"/>' for i, p in enumerate((1094, 3789), 3))
        ext = etree.fromstring(
            '<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:ext uri="{27BBF7A9-308A-43DC-89C8-2F10F3537804}">'
            f'<p15:sldGuideLst xmlns:p15="{ns}">{guides}</p15:sldGuideLst></p:ext></p:extLst>'
        )
        prs.slide_master._element.append(ext)

    return make_deck(build, name="guided.pptx")


@pytest.fixture
def examples_with_three_accent1_rects(make_deck):
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Cm

    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        for i in range(3):
            shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1 + i * 5), Cm(5), Cm(4), Cm(3))
            shp.fill.solid()
            shp.fill.fore_color.theme_color = 5  # ACCENT_1

    return make_deck(build, name="examples.pptx")


@pytest.fixture
def slide_with_box():
    """(slide, textbox): a blank slide and a factory for word-wrapped text boxes on it.

    `textbox(width_cm, text, size, family=None, height_cm=5)` sets `text` through python-pptx
    ("\\n" starts a paragraph, "\\v" is a line break) and gives every run `size` pt and `family`.
    """
    from pptx.util import Cm, Pt

    prs = Presentation()
    slide = prs.slides.add_slide(layout_named(prs, "Blank"))

    def textbox(width_cm, text, size, family=None, height_cm=5):
        box = slide.shapes.add_textbox(Cm(1), Cm(1), Cm(width_cm), Cm(height_cm))
        box.text_frame.word_wrap = True
        box.text_frame.text = text
        for paragraph in box.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(size)
                if family:
                    run.font.name = family
        return box

    return slide, textbox


@pytest.fixture
def deck_dir_from_default(tmp_path, default_pptx):
    from pptx_skills.deckdir import init_deck

    deck_dir = tmp_path / "decks/d"
    init_deck(deck_dir, default_pptx)
    return deck_dir


def finding(report, rule, shape=None):
    """The first finding of `rule` (optionally on the shape named `shape`); None when absent."""
    return next((f for f in report["findings"] if f["rule"] == rule and shape in (None, f["shape"])), None)
