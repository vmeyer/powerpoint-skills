"""`pptx exec`: agent-written python-pptx scripts with the safety of apply."""

import json
import textwrap

import pytest
from pptx.oxml.ns import qn

from conftest import assert_sections_consistent, sectioned_deck
from pptx_skills.changeset import apply_changeset
from pptx_skills.cli import main
from pptx_skills.deckdir import init_deck, load_deck
from pptx_skills.errors import PptxSkillsError
from pptx_skills.opc import file_hash, open_presentation
from pptx_skills.scripting import run_script


def script(tmp_path, body, name="big-number.py"):
    path = tmp_path / name
    path.write_text(textwrap.dedent(body), encoding="utf-8")
    return path


@pytest.fixture
def work(tmp_path):
    """A deck folder without master and with one slide; returns the working copy."""
    init_deck(tmp_path / "decks/s", None)
    deck = tmp_path / "decks/s/s.pptx"
    apply_changeset(deck, {"label": "first", "ops": [{"op": "add_slide", "layout": "Title Only"}]})
    return deck


def _state(deck):
    folder = deck.parent
    return file_hash(deck), (folder / "deck.json").read_text(), sorted(p.name for p in (folder / "changes").iterdir())


BIG_NUMBER = """
    from pptx.enum.shapes import MSO_SHAPE

    slide = add_slide("Blank")
    band = plain(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, cm(8)))
    band.name = "band"
    band.fill.gradient()
    band.fill.gradient_angle = 90
    theme_color(band.fill.gradient_stops[0].color, "accent1")
    theme_color(band.fill.gradient_stops[1].color, "accent1-50")
    number = plain(slide.shapes.add_textbox(cm(2), cm(1), cm(20), cm(6)))
    number.name = "number"
    run = number.text_frame.paragraphs[0].add_run()
    run.text = "-40 %"
    run.font.size = pt(96)
    theme_color(run.font, "#FFFFFF")
    marker = plain(slide.shapes.add_shape(MSO_SHAPE.OVAL, cm(25), cm(10), cm(3), cm(3)))
    theme_color(marker.fill, "accent2+40")
    theme_color(marker.line, "none")
    assert deck_dir.name == "s" and brief["slide_size_cm"][0] > 33
    assert slide_by_id(slide.slide_id) is slide and layout("Blank").name == "Blank"
"""


def test_a_script_changes_the_deck_with_snapshot_and_log(work, tmp_path):
    before = work.read_bytes()
    path = script(tmp_path, BIG_NUMBER)
    result = run_script(work, path)
    prs = open_presentation(work)
    slide = prs.slides[1]
    assert result["seq"] == 2 and result["touched"] == [slide.slide_id]
    shapes = {s.name: s for s in slide.shapes}
    assert {"Date Placeholder 1", "band", "number"} <= set(shapes)  # the helper copies footer, date and number
    band = shapes["band"]
    assert band._element.find(qn("p:style")) is None
    assert [stop.find(qn("a:schemeClr")).get("val") for stop in band._element.iter(qn("a:gs"))] == ["accent1", "accent1"]
    assert str(shapes["number"].text_frame.paragraphs[0].runs[0].font.color.rgb) == "FFFFFF"
    folder = work.parent
    assert (folder / "changes/002-big-number.py").read_text() == path.read_text()
    assert (folder / "snapshots/002.pptx").read_bytes() == before
    assert load_deck(folder)["last_hash"] == file_hash(work) and load_deck(folder)["last_touched"] == [slide.slide_id]


@pytest.mark.parametrize("body,needle", [
    ("x = 1\nraise ValueError('boom')\n", "line 3: ValueError: boom"),
    ("prs.save('elsewhere.pptx')\n", "must not save"),
    ("slide_by_id(9999)\n", "no slide with id 9999"),
    ("def broken(:\n", "SyntaxError"),
    ("add_slide('Nope')\n", "unknown layout"),
])
def test_a_failing_script_changes_nothing(work, tmp_path, body, needle):
    state = _state(work)
    with pytest.raises(PptxSkillsError, match="nothing was changed") as exc:
        run_script(work, script(tmp_path, "add_slide('Blank')\n" + body, "fails.py"))
    assert needle in str(exc.value) and "fails.py" in str(exc.value)
    assert _state(work) == state
    assert not list(work.parent.glob(".*.tmp")) and not (tmp_path / "elsewhere.pptx").exists()


def test_a_script_that_breaks_the_structure_is_rolled_back(work, tmp_path):
    state = _state(work)
    body = """
        from lxml import etree
        pic = etree.fromstring(
            '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<p:nvPicPr><p:cNvPr id="99" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
            '<p:blipFill><a:blip r:embed="rId77"/></p:blipFill><p:spPr/></p:pic>')
        prs.slides[0].shapes._spTree.append(pic)
    """
    with pytest.raises(PptxSkillsError, match="relationships.*nothing was changed|nothing was changed.*rId77"):
        run_script(work, script(tmp_path, body, "dangling.py"))
    assert _state(work) == state


def test_only_slides_whose_xml_changed_count_as_touched(work, tmp_path):
    apply_changeset(work, {"ops": [{"op": "add_slide", "layout": "Title Only"}, {"op": "add_slide", "layout": "Title Only"}]})
    prs = open_presentation(work)
    ids = [s.slide_id for s in prs.slides]
    result = run_script(work, script(tmp_path, "prs.slides[1].shapes.title.text = 'Changed'\n", "retitle.py"))
    assert result["touched"] == [ids[1]]
    result = run_script(work, script(tmp_path, "add_slide('Blank', after=prs.slides[0])\n", "insert.py"))
    new = open_presentation(work).slides[1].slide_id
    assert result["touched"] == [new] and new > max(ids)  # the slides after it got new numbers, not new content


def test_slides_added_without_the_helper_are_normalised(tmp_path, make_deck):
    init_deck(tmp_path / "decks/sec", sectioned_deck(make_deck), keep_slides=True)
    deck = tmp_path / "decks/sec/sec.pptx"
    last = open_presentation(deck).slides[2].slide_id
    apply_changeset(deck, {"ops": [{"op": "delete_slide", "slide": 3}]})  # its id is never used again (R26)
    body = """
        new = prs.slides.add_slide(layout("Title Only"))  # python-pptx alone would reuse the deleted id
        new.shapes.title.text = "Added directly"
    """
    result = run_script(deck, script(tmp_path, body, "direct.py"))
    prs = open_presentation(deck)
    added = prs.slides[2]
    assert added.slide_id > last and result["touched"] == [added.slide_id]
    assert load_deck(tmp_path / "decks/sec")["max_slide_id"] == added.slide_id
    assert_sections_consistent(prs)
    names = [s.part.partname for s in prs.slides]
    assert len(set(names)) == len(names)


def test_exec_writes_only_working_copies(default_pptx, potx_file, tmp_path, make_deck):
    body = script(tmp_path, "add_slide('Blank')\n", "add.py")
    with pytest.raises(PptxSkillsError, match="--outside-deck"):
        run_script(default_pptx, body)
    with pytest.raises(PptxSkillsError, match=r"\.potx"):
        run_script(potx_file, body, outside_deck=True)
    result = run_script(default_pptx, body, outside_deck=True)
    assert result["seq"] is None and len(open_presentation(default_pptx).slides) == 1
    assert any("no snapshot" in m for m in result["messages"])


def test_cli_exec(work, tmp_path, capsys):
    path = script(tmp_path, BIG_NUMBER)
    assert main(["exec", str(work), str(path)]) == 0
    out = capsys.readouterr().out
    assert "Ran big-number.py as change 002" in out and "touched slide ids" in out
    assert main(["exec", str(work), str(tmp_path / "missing.py"), "--json"]) == 2
    assert "missing.py" in capsys.readouterr().err
    assert main(["exec", str(work), str(script(tmp_path, "pass\n", "noop.py")), "--json"]) == 0
    assert json.loads(capsys.readouterr().out)["touched"] == []


# --- slides deleted or moved by a script (R12) -----------------------------------------------------

RAW_DELETE = """
    first = prs.slides._sldIdLst[0]
    prs.part.drop_rel(first.rId)
    prs.slides._sldIdLst.remove(first)
"""
RAW_MOVE = """
    ids = prs.slides._sldIdLst
    first = ids[0]
    ids.remove(first)
    ids.append(first)
"""


@pytest.fixture
def sectioned_work(tmp_path, make_deck):
    init_deck(tmp_path / "decks/sec", sectioned_deck(make_deck), keep_slides=True)
    return tmp_path / "decks/sec/sec.pptx"


@pytest.mark.parametrize("body", [RAW_DELETE, RAW_MOVE, "move_slide(prs.slides[2], 1)\n" + textwrap.dedent(RAW_MOVE)])
def test_a_script_that_deletes_or_moves_slides_itself_is_refused(sectioned_work, tmp_path, body):
    state = _state_without_changes(sectioned_work)
    with pytest.raises(PptxSkillsError, match="nothing was changed") as exc:
        run_script(sectioned_work, script(tmp_path, body, "raw.py"))
    assert "delete_slide" in str(exc.value) and "move_slide" in str(exc.value)
    assert _state_without_changes(sectioned_work) == state


def _state_without_changes(deck):
    return file_hash(deck), (deck.parent / "deck.json").read_text()


def test_the_delete_and_move_helpers_keep_sections_links_and_numbers(sectioned_work, tmp_path):
    before = [s.slide_id for s in open_presentation(sectioned_work).slides]
    body = """
        delete_slide(prs.slides[0])
        move_slide(slide_by_id(%d), 1)
        added = add_slide("Blank", after=1)
        delete_slide(added)
        assert len(prs.slides) == 2
    """ % before[2]
    result = run_script(sectioned_work, script(tmp_path, body, "helpers.py"))
    prs = open_presentation(sectioned_work)
    assert [s.slide_id for s in prs.slides] == [before[2], before[1]]
    assert_sections_consistent(prs)
    assert sorted(result["touched"]) == []  # a move or a delete changes no slide's content
    names = [str(s.part.partname) for s in prs.slides]
    assert names == ["/ppt/slides/slide1.xml", "/ppt/slides/slide2.xml"]


# --- colors, output and master changes --------------------------------------------------------

def test_theme_color_replaces_the_template_gradients_modifiers(work, tmp_path):
    body = """
        from pptx.enum.shapes import MSO_SHAPE
        slide = add_slide("Blank")
        band = plain(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, cm(10), cm(5)))
        band.name = "band"
        band.fill.gradient()
        theme_color(band.fill.gradient_stops[0].color, "accent1")
        theme_color(band.fill.gradient_stops[1].color, "accent1-50")
    """
    run_script(work, script(tmp_path, body, "gradient.py"))
    band = next(s for s in open_presentation(work).slides[1].shapes if s.name == "band")
    stops = [gs.find(qn("a:schemeClr")) for gs in band._element.iter(qn("a:gs"))]
    assert [[etree_name(c) for c in stop] for stop in stops] == [[], ["lumMod"]]


def etree_name(element) -> str:
    return element.tag.split("}")[1]


def test_print_in_a_script_goes_to_stderr(work, tmp_path, capsys):
    path = script(tmp_path, "print('hello from the script')\nadd_slide('Blank')\n", "talks.py")
    assert main(["exec", str(work), str(path), "--json"]) == 0
    captured = capsys.readouterr()
    assert json.loads(captured.out)["seq"] == 2 and "hello from the script" in captured.err


def test_a_master_or_theme_change_touches_every_slide(work, tmp_path):
    apply_changeset(work, {"ops": [{"op": "add_slide", "layout": "Blank"}]})
    ids = [s.slide_id for s in open_presentation(work).slides]
    body = """
        prs.slide_master.background.fill.solid()
        theme_color(prs.slide_master.background.fill, "lt2")
    """
    result = run_script(work, script(tmp_path, body, "background.py"))
    assert sorted(result["touched"]) == sorted(ids)
    assert any("master" in m for m in result["messages"])
