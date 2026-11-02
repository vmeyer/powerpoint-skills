import pytest
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

from conftest import apply, finding, layout_named
from pptx_skills.checks import run_checks
from pptx_skills.deckdir import init_deck, load_deck
from pptx_skills.errors import PptxSkillsError
from pptx_skills.inventory import inventory
from pptx_skills.opc import open_presentation


def table_item(deck, name="tbl"):
    shapes = inventory(open_presentation(deck))["slides"][0]["shapes"]
    return next(s for s in shapes if s["name"] == name)


def frame_named(deck, name="tbl"):
    return next(s for s in open_presentation(deck).slides[0].shapes if s.name == name)


@pytest.fixture
def deck_with_table(make_deck):
    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        slide.shapes.add_table(2, 2, Cm(2), Cm(2), Cm(10), Cm(3)).name = "tbl"

    return make_deck(build, name="table.pptx")


def _example_table(prs):
    slide = prs.slides.add_slide(layout_named(prs, "Blank"))
    frame = slide.shapes.add_table(2, 2, Cm(2), Cm(2), Cm(10), Cm(3))
    frame.name = "tbl"
    table = frame.table
    for r, row in enumerate([["Head A", "Head B"], ["Cell one", "Cell two"]]):
        for c, text in enumerate(row):
            table.cell(r, c).text = text
    run = table.cell(1, 1).text_frame.paragraphs[0].runs[0]
    run.font.bold, run.font.size = True, Pt(13)
    return slide


@pytest.fixture
def cloned_table(tmp_path, default_pptx, make_deck):
    """A deck folder whose working copy holds a table cloned from the examples."""
    examples = make_deck(_example_table, name="ex.pptx")
    folder = tmp_path / "decks/t"
    init_deck(folder, default_pptx, examples)
    deck = folder / "t.pptx"
    apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 1}}])
    return folder, deck


def test_add_table(make_deck):
    deck = make_deck(lambda prs: prs.slides.add_slide(layout_named(prs, "Blank")))
    apply(deck, [{"op": "add_table", "slide": 1, "name": "tbl", "box": [2, 3, 20, 6],
                  "rows": [["A", "B"], ["1", "2"], ["3", "4"]], "col_widths": [8, 12], "fills": {"0,0": "accent1"}}])
    item = table_item(deck)
    assert item["table"] == [["A", "B"], ["1", "2"], ["3", "4"]]
    assert item["kind"] == "table" and item["box"] == [2, 3, 20, 6]
    frame = frame_named(deck)
    assert [round(c.width / 360000) for c in frame.table.columns] == [8, 12]
    assert frame.table.cell(0, 0).fill.fore_color.theme_color is not None
    assert frame.table.first_row and not frame.table.horz_banding


def test_add_table_options_and_errors(make_deck):
    deck = make_deck(lambda prs: prs.slides.add_slide(layout_named(prs, "Blank")))
    guid = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    apply(deck, [{"op": "add_table", "slide": 1, "name": "tbl", "box": [2, 3, 20, 6], "rows": [["**Bold** plain", "x"]],
                  "header": False, "style": guid, "font": {"size": 12}}])
    frame = frame_named(deck)
    assert not frame.table.first_row
    assert frame._element.graphic.graphicData.tbl.tblPr.find(qn("a:tableStyleId")).text == guid
    runs = frame.table.cell(0, 0).text_frame.paragraphs[0].runs
    assert [(r.text, r.font.bold, r.font.size.pt) for r in runs] == [("Bold", True, 12), (" plain", False, 12)]
    for bad, message in (
        ({"rows": [["a", "b"], ["c"]]}, "columns"),
        ({"rows": [["a", "b"]], "col_widths": [5]}, "col_widths"),
        ({"rows": [["a", "b"]], "fills": {"3,0": "accent1"}}, "outside"),
        ({"rows": [["a", "b"]], "name": "tbl"}, "already exists"),
    ):
        with pytest.raises(PptxSkillsError, match=message):
            apply(deck, [{"op": "add_table", "slide": 1, "name": "other", "box": [2, 9, 10, 2], **bad}])


CUSTOM = "{11111111-2222-3333-4444-555555555555}"
OTHER = "{66666666-7777-8888-9999-AAAAAAAAAAAA}"


@pytest.fixture
def styled_deck(make_deck):
    """A deck whose tableStyles part defines two custom styles and makes the first one the default."""
    xml = (
        '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="%s">'
        '<a:tblStyle styleId="%s" styleName="Corporate lined"/><a:tblStyle styleId="%s" styleName="Corporate filled"/></a:tblStyleLst>'
    ) % (CUSTOM, CUSTOM, OTHER)

    def build(prs):
        prs.slides.add_slide(layout_named(prs, "Blank"))
        part = next(r.target_part for r in prs.part.rels.values() if r.reltype.endswith("/tableStyles"))
        part._blob = xml.encode()

    return make_deck(build, name="styled.pptx")


def style_of(deck, name="tbl"):
    return frame_named(deck, name)._element.graphic.graphicData.tbl.tblPr.find(qn("a:tableStyleId")).text


def test_table_style_default_name_and_unknown(styled_deck):
    from pptx_skills.brief import table_style_catalog

    base = {"op": "add_table", "slide": 1, "box": [2, 3, 10, 2], "rows": [["a", "b"]]}
    apply(styled_deck, [{**base, "name": "tbl"}, {**base, "name": "t2", "box": [2, 7, 10, 2], "style": "corporate FILLED"}])
    assert style_of(styled_deck) == CUSTOM and style_of(styled_deck, "t2") == OTHER
    assert table_style_catalog(open_presentation(styled_deck)) == [
        {"name": "Corporate lined", "id": CUSTOM, "default": True}, {"name": "Corporate filled", "id": OTHER, "default": False}]
    with pytest.raises(PptxSkillsError, match='unknown table style "nope".*Corporate lined.*Corporate filled'):
        apply(styled_deck, [{**base, "name": "t3", "style": "nope"}])


def test_without_custom_styles_the_builtin_default_stays(make_deck):
    deck = make_deck(lambda prs: prs.slides.add_slide(layout_named(prs, "Blank")))
    apply(deck, [{"op": "add_table", "slide": 1, "name": "tbl", "box": [2, 3, 10, 2], "rows": [["a"]]}])
    assert style_of(deck) == "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"


def test_brief_lists_table_styles(styled_deck, tmp_path):
    from pptx_skills.brief import _markdown, build_brief

    brief = build_brief(open_presentation(styled_deck))
    assert [e["name"] for e in brief["table_style_catalog"]] == ["Corporate lined", "Corporate filled"]
    assert f"- Corporate lined {CUSTOM} (default)" in _markdown(brief, "")


def test_bold_markers_determine_the_whole_cell(make_deck):
    deck = make_deck(lambda prs: prs.slides.add_slide(layout_named(prs, "Blank")))
    apply(deck, [{"op": "add_table", "slide": 1, "name": "tbl", "box": [2, 3, 10, 2], "col_widths": [3.333, 4.4449],
                  "rows": [["**Bold** plain", "x"]], "font": {"bold": False}}])
    frame = frame_named(deck)
    assert frame.width == sum(c.width for c in frame.table.columns)
    runs = frame.table.cell(0, 0).text_frame.paragraphs[0].runs
    assert [r.font.bold for r in runs] == [True, False]
    apply(deck, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": [["**Bold** plain", "x"]]}])
    assert [r.font.bold for r in frame_named(deck).table.cell(0, 0).text_frame.paragraphs[0].runs] == [True, False]


def test_set_table_refuses_to_copy_or_cut_vertical_merges(make_deck):
    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        slide.shapes.add_table(3, 2, Cm(2), Cm(2), Cm(10), Cm(3)).name = "tbl"
        slide.shapes[0].table.cell(1, 0).merge(slide.shapes[0].table.cell(2, 0))

    deck = make_deck(build)
    with pytest.raises(PptxSkillsError, match="vertically merged.*adding rows"):
        apply(deck, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": [["a", "b"]] * 4}])
    with pytest.raises(PptxSkillsError, match="cut a vertically merged cell"):
        apply(deck, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": [["a", "b"]] * 2}])
    apply(deck, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": [["a", "b"], ["c", "d"], ["", "e"]]}])
    assert table_item(deck)["table"][1][0] == "c"


def test_set_table_grows_and_shrinks(cloned_table):
    folder, deck = cloned_table
    rows = [["H1", "H2"], ["a", "b"], ["c", "d"], ["e", "f"]]
    apply(deck, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": rows}])
    assert table_item(deck)["table"] == rows
    table = frame_named(deck).table
    for r in (1, 2, 3):  # rows copied from the last one keep its formatting; plain text takes no weight (R13')
        run = table.cell(r, 1).text_frame.paragraphs[0].runs[0]
        assert run.font.bold is None and run.font.size.pt == 13
    assert frame_named(deck).height == sum(int(tr.get("h")) for tr in table._tbl.tr_lst)
    apply(deck, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": [["only", "row"]]}])
    assert table_item(deck)["table"] == [["only", "row"]]


def test_copied_rows_get_no_duplicate_row_ids(cloned_table):
    from lxml import etree

    folder, deck = cloned_table
    prs = open_presentation(deck)
    last = next(s for s in prs.slides[0].shapes if s.name == "tbl").table._tbl.tr_lst[-1]
    ext = etree.fromstring(
        '<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext uri="{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}">'
        '<a16:rowId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="42"/></a:ext></a:extLst>'
    )
    last.append(ext)
    prs.save(str(deck))
    apply(deck, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": [["a", "b"]] * 4}])
    table = frame_named(deck).table
    assert [len(tr.findall(qn("a:extLst"))) for tr in table._tbl.tr_lst] == [0, 1, 0, 0]  # only the original row keeps its id


def test_header_stays_unbolded_by_plain_runs(cloned_table):
    folder, deck = cloned_table
    apply(deck, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": [["New head", "**B**"], ["x", "y"]]}])
    table = frame_named(deck).table
    assert table.cell(0, 0).text_frame.paragraphs[0].runs[0].font.bold is None  # inherits the table style
    assert table.cell(0, 1).text_frame.paragraphs[0].runs[0].font.bold is True


def test_column_mismatch(deck_with_table):
    with pytest.raises(PptxSkillsError, match="columns"):
        apply(deck_with_table, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": [["a", "b", "c"]]}])
    with pytest.raises(PptxSkillsError, match="not a table"):
        apply(deck_with_table, [{"op": "add_table", "slide": 1, "name": "t2", "box": [1, 1, 5, 2], "rows": [["a"]]},
                                {"op": "set_table", "slide": 1, "target": {"name": "t2"}, "rows": [["a", "b"]]}][:1]
              + [{"op": "add_shape", "slide": 1, "name": "box", "shape": "rect", "box": [1, 9, 3, 2]},
                 {"op": "set_table", "slide": 1, "target": {"name": "box"}, "rows": [["a"]]}])


def test_cloned_table_cells_are_recorded_and_flagged(cloned_table):
    folder, deck = cloned_table
    slide = open_presentation(deck).slides[0]
    frame = frame_named(deck)
    recorded = load_deck(folder)["cloned_texts"][str(slide.slide_id)]
    key = lambda r, c: f"{frame.shape_id}:r{r}c{c}"  # noqa: E731
    assert recorded == {key(0, 0): "Head A", key(0, 1): "Head B", key(1, 0): "Cell one", key(1, 1): "Cell two"}
    flagged = [f for f in run_checks(deck)["findings"] if f["rule"] == "leftover-text"]
    assert len(flagged) == 4 and all(f["shape"] == "tbl" for f in flagged)
    apply(deck, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": [["Region", "Revenue"], ["North", "4 TODO"]]}])
    flagged = [f["message"] for f in run_checks(deck)["findings"] if f["rule"] == "leftover-text"]
    assert len(flagged) == 1 and "row 2, column 2" in flagged[0] and "TODO" in flagged[0]
    apply(deck, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": [["Region", "Revenue"], ["North", "4 M"]]}])
    assert finding(run_checks(deck), "leftover-text") is None


def test_lorem_in_a_cell_is_flagged_without_deck_folder(deck_with_table):
    apply(deck_with_table, [{"op": "set_table", "slide": 1, "target": {"name": "tbl"}, "rows": [["Lorem ipsum", "ok"], ["x", "y"]]}])
    assert finding(run_checks(deck_with_table), "leftover-text", "tbl")
