import json

import pytest
from lxml import etree
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.util import Cm

from conftest import layout_named
from pptx_skills.changeset import apply_changeset
from pptx_skills.checks import run_checks
from pptx_skills.cli import main
from pptx_skills.deckdir import load_deck
from pptx_skills.errors import PptxSkillsError
from pptx_skills.opc import file_hash, open_presentation
from pptx_skills import slides

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def test_add_and_move_slides(deck_dir_from_default):
    deck = deck_dir_from_default / "d.pptx"
    res = apply_changeset(deck, {"label": "two", "ops": [
        {"op": "add_slide", "layout": "Title Only"}, {"op": "add_slide", "layout": "Blank"},
        {"op": "move_slide", "slide": "$last", "to": 1}]})
    prs = open_presentation(deck)
    assert [s.slide_layout.name for s in prs.slides] == ["Blank", "Title Only"]
    assert res["seq"] == 1 and (deck_dir_from_default / "changes/001-two.json").exists()
    assert (deck_dir_from_default / "snapshots/001.pptx").exists()
    assert load_deck(deck_dir_from_default)["last_hash"] == file_hash(deck)
    assert sorted(res["touched"]) == sorted(s.slide_id for s in prs.slides)


def test_atomic_on_error(deck_dir_from_default):
    deck = deck_dir_from_default / "d.pptx"; before = file_hash(deck)
    with pytest.raises(PptxSkillsError, match="op 2 .delete_slide.: ") as exc:
        apply_changeset(deck, {"label": "bad", "ops": [{"op": "add_slide", "layout": "Title Only"}, {"op": "delete_slide", "slide": 9}]})
    assert str(exc.value).startswith("op 2 (delete_slide): ")
    assert file_hash(deck) == before
    assert not (deck_dir_from_default / "changes").exists() or not list((deck_dir_from_default / "changes").iterdir())
    assert not (deck_dir_from_default / "snapshots").exists() or not list((deck_dir_from_default / "snapshots").iterdir())
    assert load_deck(deck_dir_from_default)["seq"] == 0
    assert not list(deck_dir_from_default.glob(".*.tmp"))


def test_schema_rejects_unknown_op(deck_dir_from_default):
    with pytest.raises(PptxSkillsError, match="op 1.*frobnicate"):
        apply_changeset(deck_dir_from_default / "d.pptx", {"label": "x", "ops": [{"op": "frobnicate"}]})


def test_schema_rejects_bad_fields(deck_dir_from_default):
    deck = deck_dir_from_default / "d.pptx"
    with pytest.raises(PptxSkillsError, match="op 1 .add_slide."):
        apply_changeset(deck, {"ops": [{"op": "add_slide", "layout": "Blank", "extra": 1}]})
    with pytest.raises(PptxSkillsError, match="op 2 .move_slide."):
        apply_changeset(deck, {"ops": [{"op": "add_slide", "layout": "Blank"}, {"op": "move_slide", "slide": "x", "to": 1}]})
    with pytest.raises(PptxSkillsError, match="at least one op"):
        apply_changeset(deck, {"ops": []})


def test_unknown_layout_lists_names(deck_dir_from_default):
    with pytest.raises(PptxSkillsError, match="Title Only"):
        apply_changeset(deck_dir_from_default / "d.pptx", {"label": "x", "ops": [{"op": "add_slide", "layout": "Nope"}]})


def test_apply_after_external_edit_reports_and_applies(deck_dir_from_default):
    deck = deck_dir_from_default / "d.pptx"
    prs = open_presentation(deck); prs.slides.add_slide(prs.slide_layouts[6]); prs.save(deck)   # hand edit
    (deck.parent / f"~${deck.name}").write_bytes(b"x")                                         # open in PowerPoint
    res = apply_changeset(deck, {"label": "more", "ops": [{"op": "add_slide", "layout": "Blank"}]})
    assert len(open_presentation(deck).slides) == 2
    assert any("changed outside" in m for m in res["messages"]) and any("open in PowerPoint" in m for m in res["messages"])


def test_label_is_sanitised(deck_dir_from_default):
    deck = deck_dir_from_default / "d.pptx"
    apply_changeset(deck, {"label": "Äh! Neue Folie: " + "x" * 60, "ops": [{"op": "add_slide", "layout": "Blank"}]})
    apply_changeset(deck, {"ops": [{"op": "add_slide", "layout": "Blank"}]})
    names = sorted(p.name for p in (deck_dir_from_default / "changes").iterdir())
    assert names[1] == "002-changes.json"
    stem = names[0][4:-5]
    assert len(stem) <= 40 and stem.startswith("h-neue-folie-xxx") and set(stem) <= set("abcdefghijklmnopqrstuvwxyz0123456789-")


def test_outside_deck_folder_applies_without_log(tmp_path, default_pptx):
    res = apply_changeset(default_pptx, {"ops": [{"op": "add_slide", "layout": "Blank"}]})
    assert len(open_presentation(default_pptx).slides) == 1
    assert res["seq"] is None and any("not inside a deck folder" in m for m in res["messages"])
    assert not (tmp_path / "changes").exists()


def test_slide_refs_and_delete(deck_dir_from_default):
    deck = deck_dir_from_default / "d.pptx"
    apply_changeset(deck, {"ops": [{"op": "add_slide", "layout": "Title Only"}, {"op": "add_slide", "layout": "Blank"},
                                   {"op": "add_slide", "layout": "Title and Content", "after": 1}]})
    prs = open_presentation(deck)
    assert [s.slide_layout.name for s in prs.slides] == ["Title Only", "Title and Content", "Blank"]
    ident = prs.slides[1].slide_id
    res = apply_changeset(deck, {"ops": [{"op": "delete_slide", "slide": f"id:{ident}"}]})
    assert [s.slide_layout.name for s in open_presentation(deck).slides] == ["Title Only", "Blank"]
    assert ident not in res["touched"]
    with pytest.raises(PptxSkillsError, match=r"op 1 .delete_slide.: no slide with id 999"):
        apply_changeset(deck, {"ops": [{"op": "delete_slide", "slide": "id:999"}]})
    with pytest.raises(PptxSkillsError, match=r'"\$last"'):
        apply_changeset(deck, {"ops": [{"op": "delete_slide", "slide": "$last"}]})
    with pytest.raises(PptxSkillsError, match="position 5"):
        apply_changeset(deck, {"ops": [{"op": "move_slide", "slide": 1, "to": 5}]})


def test_cli_apply(deck_dir_from_default, tmp_path, capsys):
    cs = tmp_path / "cs.json"
    cs.write_text(json.dumps({"label": "cli", "ops": [{"op": "add_slide", "layout": "Blank"}]}))
    assert main(["apply", str(deck_dir_from_default / "d.pptx"), str(cs), "--json"]) == 0
    assert json.loads(capsys.readouterr().out)["seq"] == 1
    cs.write_text(json.dumps({"ops": [{"op": "nope"}]}))
    assert main(["apply", str(deck_dir_from_default / "d.pptx"), str(cs)]) == 2
    assert "op 1 (nope)" in capsys.readouterr().err


# --- sections (R12) ---------------------------------------------------------------------------

def _sectioned(make_deck):
    """Three slides in sections A [s1, s2] and B [s3]."""
    def build(prs):
        for _ in range(3):
            prs.slides.add_slide(layout_named(prs, "Blank"))
        ids = [s.id for s in prs.slides._sldIdLst]
        sections = "".join(
            f'<p14:section name="{n}" id="{{0000000{i}-0000-0000-0000-000000000000}}"><p14:sldIdLst>'
            + "".join(f'<p14:sldId id="{x}"/>' for x in members) + "</p14:sldIdLst></p14:section>"
            for i, (n, members) in enumerate((("A", ids[:2]), ("B", ids[2:])), 1)
        )
        ext = etree.fromstring(
            '<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">'
            f'<p14:sectionLst xmlns:p14="{P14}">{sections}</p14:sectionLst></p:ext></p:extLst>'
        )
        prs.part._element.append(ext)
    return make_deck(build, name="sectioned.pptx")


def _sections(prs):
    return [[int(r.get("id")) for r in sec.iter(f"{{{P14}}}sldId")] for sec in prs.part._element.iter(f"{{{P14}}}section")]


def _consistent(prs):
    flat = [i for sec in _sections(prs) for i in sec]
    assert flat == [s.slide_id for s in prs.slides]  # exactly one section each, same order


def test_sections_follow_added_and_moved_slides(make_deck, tmp_path):
    path = _sectioned(make_deck)
    prs = open_presentation(path)
    a1, a2, b1 = (s.slide_id for s in prs.slides)
    new = slides.add_slide(prs, "Blank", after=prs.slides[0])  # after s1: section A
    assert _sections(prs) == [[a1, new.slide_id, a2], [b1]]; _consistent(prs)
    end = slides.add_slide(prs, "Blank")  # appended after s3: section B
    assert _sections(prs) == [[a1, new.slide_id, a2], [b1, end.slide_id]]; _consistent(prs)
    slides.move_slide(prs, end, 1)  # first position: first section
    assert _sections(prs)[0][0] == end.slide_id; _consistent(prs)
    slides.move_slide(prs, prs.slides[0], 5)  # last position: last section
    assert _sections(prs)[1][-1] == end.slide_id; _consistent(prs)
    slides.move_slide(prs, prs.slides[4], 3)  # into the middle of A
    assert end.slide_id in _sections(prs)[0]; _consistent(prs)
    out = tmp_path / "out.pptx"; prs.save(out)
    _consistent(open_presentation(out))


def test_section_for_slide_added_to_emptied_deck(make_deck):
    prs = open_presentation(_sectioned(make_deck))
    for s in list(prs.slides):
        slides.delete_slide(prs, s)
    new = slides.add_slide(prs, "Blank")
    assert _sections(prs) == [[], [new.slide_id]]
    _consistent(prs)


# --- failure handling -------------------------------------------------------------------------

def _state(deck_dir):
    deck = deck_dir / "d.pptx"
    names = lambda n: sorted(p.name for p in (deck_dir / n).iterdir()) if (deck_dir / n).is_dir() else None
    return (file_hash(deck), (deck_dir / "deck.json").read_bytes(), names("changes"), names("snapshots"),
            sorted(p.name for p in deck_dir.iterdir()))


def test_missing_deck_is_a_clean_error(tmp_path, capsys):
    cs = tmp_path / "cs.json"
    cs.write_text(json.dumps({"ops": [{"op": "add_slide", "layout": "Blank"}]}))
    with pytest.raises(PptxSkillsError, match="does not exist"):
        apply_changeset(tmp_path / "missing.pptx", json.loads(cs.read_text()))
    assert main(["apply", str(tmp_path / "missing.pptx"), str(cs)]) == 2
    assert capsys.readouterr().err.startswith("error:")


def test_failed_replace_leaves_everything_unchanged(deck_dir_from_default, monkeypatch):
    apply_changeset(deck_dir_from_default / "d.pptx", {"label": "first", "ops": [{"op": "add_slide", "layout": "Blank"}]})
    state = _state(deck_dir_from_default)

    def boom(*a, **k):
        raise OSError("disk says no")

    monkeypatch.setattr("pptx_skills.changeset.os.replace", boom)
    with pytest.raises(PptxSkillsError, match="nothing was changed"):
        apply_changeset(deck_dir_from_default / "d.pptx", {"ops": [{"op": "add_slide", "layout": "Blank"}]})
    assert _state(deck_dir_from_default) == state


def test_failure_inside_record_apply_rolls_back(deck_dir_from_default, monkeypatch):
    state = _state(deck_dir_from_default)

    def boom(*a, **k):
        raise OSError("disk says no")

    monkeypatch.setattr("pptx_skills.deckdir._save_deck", boom)  # runs after log and snapshot were written
    with pytest.raises(PptxSkillsError, match="nothing was changed"):
        apply_changeset(deck_dir_from_default / "d.pptx", {"ops": [{"op": "add_slide", "layout": "Blank"}]})
    assert _state(deck_dir_from_default) == state


def _footer_layout_deck(make_deck):
    """Default template whose "Title Only" layout has a footer text besides its date and slide-number fields.

    The slide-number box is narrow and its cached field text a long prompt, as in real masters: only the
    number itself fits.
    """
    def build(prs):
        layout = layout_named(prs, "Title Only")
        kinds = {ph.placeholder_format.type: ph for ph in layout.placeholders}
        kinds[PP_PLACEHOLDER.FOOTER].text_frame.text = "|   Example Ltd"
        number = kinds[PP_PLACEHOLDER.SLIDE_NUMBER]
        number.width, number.height = Cm(1), Cm(1)
        number._element.find(".//" + qn("a:fld")).find(qn("a:t")).text = "‹Nr.›"

    return make_deck(build, name="footer.pptx")


def _fields(shape) -> list[str]:
    return [f.get("type") for f in shape._element.iter(qn("a:fld"))]


def test_add_slide_copies_date_footer_and_number(make_deck):
    deck = _footer_layout_deck(make_deck)
    apply_changeset(deck, {"ops": [
        {"op": "add_slide", "layout": "Title Only"},
        {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "Footers come from the layout"}]})
    slide = open_presentation(deck).slides[0]
    layout = {ph.placeholder_format.type: ph for ph in slide.slide_layout.placeholders}
    placed = {ph.placeholder_format.type: ph for ph in slide.placeholders}
    for kind in (PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER):
        assert kind in placed, kind
        assert placed[kind].placeholder_format.idx == layout[kind].placeholder_format.idx
        if kind != PP_PLACEHOLDER.SLIDE_NUMBER:
            assert placed[kind].text_frame.text == layout[kind].text_frame.text
    assert placed[PP_PLACEHOLDER.FOOTER].text_frame.text == "|   Example Ltd"
    assert _fields(placed[PP_PLACEHOLDER.SLIDE_NUMBER]) == ["slidenum"]
    assert placed[PP_PLACEHOLDER.SLIDE_NUMBER].text_frame.text == "1"  # cached field text, as PowerPoint stores it
    assert _fields(placed[PP_PLACEHOLDER.DATE]) == _fields(layout[PP_PLACEHOLDER.DATE]) != []
    assert len({s.shape_id for s in slide.shapes}) == len(slide.shapes)

    report = run_checks(deck, all_slides=True)
    assert not [f for f in report["findings"] if f["rule"] == "empty-placeholder"]
    assert report["summary"]["error"] == 0


def test_slide_number_text_follows_position(make_deck):
    deck = _footer_layout_deck(make_deck)
    apply_changeset(deck, {"ops": [{"op": "add_slide", "layout": "Title Only"}, {"op": "add_slide", "layout": "Title Only"},
                                   {"op": "add_slide", "layout": "Title Only", "after": 1}]})
    numbers = [next(ph for ph in s.placeholders if ph.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER).text_frame.text
               for s in open_presentation(deck).slides]
    assert numbers == ["1", "2", "3"]
