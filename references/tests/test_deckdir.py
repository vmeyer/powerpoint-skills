import json
import zipfile

import pytest

from conftest import linked_deck
from pptx_skills.brief import BLOCK_END

from pptx_skills.cli import main
from pptx_skills.deckdir import changed_externally, find_deck_dir, init_deck, load_deck, record_apply
from pptx_skills.errors import PptxSkillsError
from pptx_skills.opc import file_hash, open_presentation
from pptx_skills.slides import delete_slide


def test_init_from_potx(tmp_path, potx_file):
    d = init_deck(tmp_path / "decks/ai-gov", potx_file)
    prs = open_presentation(tmp_path / "decks/ai-gov/ai-gov.pptx")
    assert len(prs.slides) == 0 and d["slug"] == "ai-gov"
    assert (tmp_path / "decks/ai-gov/style/brief.json").exists()
    assert d["sources"]["master"] == str(potx_file.resolve())


def test_init_strips_or_keeps_slides(tmp_path, deck_with_two_slides):
    init_deck(tmp_path / "a", deck_with_two_slides)
    init_deck(tmp_path / "b", deck_with_two_slides, keep_slides=True)
    assert len(open_presentation(tmp_path / "a/a.pptx").slides) == 0
    assert len(open_presentation(tmp_path / "b/b.pptx").slides) == 2


def test_stripped_slides_are_not_saved(tmp_path, deck_with_two_slides):
    init_deck(tmp_path / "a", deck_with_two_slides)
    import zipfile

    with zipfile.ZipFile(tmp_path / "a/a.pptx") as z:
        assert not [n for n in z.namelist() if n.startswith("ppt/slides/slide")]


def test_init_refuses_existing(tmp_path, potx_file):
    init_deck(tmp_path / "x", potx_file)
    with pytest.raises(PptxSkillsError, match="--reuse or --replace"):
        init_deck(tmp_path / "x", potx_file)


def test_init_never_writes_source(tmp_path, deck_with_two_slides):
    before = file_hash(deck_with_two_slides)
    init_deck(tmp_path / "y", deck_with_two_slides, keep_slides=True)
    assert file_hash(deck_with_two_slides) == before


def test_changed_externally(tmp_path, potx_file):
    init_deck(tmp_path / "z", potx_file)
    assert not changed_externally(tmp_path / "z")
    with open(tmp_path / "z/z.pptx", "ab") as f:
        f.write(b" ")
    assert changed_externally(tmp_path / "z")


def test_usage_errors(tmp_path, potx_file):
    with pytest.raises(PptxSkillsError, match="--from"):
        init_deck(tmp_path / "n", None)


def test_style_file_rules_and_invalid_json(tmp_path, potx_file):
    style = tmp_path / "s.json"
    style.write_text(json.dumps({"rules": {"max_bullets": 7}}))
    init_deck(tmp_path / "r", potx_file, style_file=style)
    brief = json.loads((tmp_path / "r/style/brief.json").read_text())
    assert brief["rules"]["max_bullets"] == 7 and brief["rules"]["min_size_pt"] == 11
    style.write_text("{nope")
    with pytest.raises(PptxSkillsError, match="not valid JSON"):
        init_deck(tmp_path / "r2", potx_file, style_file=style)


def test_reuse_keeps_working_copy_and_building_blocks(tmp_path, potx_file):
    d = tmp_path / "u"
    init_deck(d, potx_file)
    md = d / "style/brief.md"
    md.write_text(md.read_text().replace(BLOCK_END, "MY BLOCK" + BLOCK_END))
    with open(d / "u.pptx", "ab") as f:
        f.write(b" ")
    before = file_hash(d / "u.pptx")
    deck = init_deck(d, None, mode="reuse")
    assert file_hash(d / "u.pptx") == before and deck["slug"] == "u"
    assert changed_externally(d) and "MY BLOCK" in md.read_text()


def test_replace_starts_over_but_leaves_other_files(tmp_path, potx_file):
    d = tmp_path / "p"
    init_deck(d, potx_file)
    (d / "changes").mkdir(exist_ok=True)
    (d / "changes/old.json").write_text("{}")
    (d / "notes.txt").write_text("keep")
    with open(d / "p.pptx", "ab") as f:
        f.write(b" ")
    init_deck(d, potx_file, mode="replace")
    assert not (d / "changes").exists() and (d / "notes.txt").read_text() == "keep"
    assert not changed_externally(d)


def test_record_apply_and_find_deck_dir(tmp_path, potx_file):
    d = tmp_path / "q"
    init_deck(d, potx_file)
    n = record_apply(d, {"ops": []}, "Fix Title!", b"snap", [256], "abc")
    assert n == 1 and (d / "changes/001-fix-title.json").is_file()
    assert (d / "snapshots/001.pptx").read_bytes() == b"snap"
    deck = load_deck(d)
    assert (deck["seq"], deck["last_hash"], deck["last_touched"]) == (1, "abc", [256])
    assert record_apply(d, {}, "x", b"", [], "def") == 2
    assert find_deck_dir(d / "q.pptx") == d.resolve()
    assert find_deck_dir(tmp_path / "other.pptx") is None


def test_delete_slide_drops_part(deck_with_two_slides, tmp_path):
    prs = open_presentation(deck_with_two_slides)
    delete_slide(prs, prs.slides[0])
    out = tmp_path / "out.pptx"
    prs.save(str(out))
    assert [s.shapes.title.text for s in open_presentation(out).slides] == ["Second slide"]
    import zipfile

    with zipfile.ZipFile(out) as z:
        assert len([n for n in z.namelist() if n.startswith("ppt/slides/slide")]) == 1


def test_cli_init_and_style(tmp_path, potx_file, capsys):
    d = tmp_path / "c"
    assert main(["init", str(d), "--from", str(potx_file), "--json"]) == 0
    assert json.loads(capsys.readouterr().out)["slug"] == "c"
    assert main(["init", str(d), "--from", str(potx_file)]) == 2
    assert "--reuse or --replace" in capsys.readouterr().err
    (d / "style/brief.json").unlink()
    assert main(["style", str(d), "--refresh"]) == 0
    assert (d / "style/brief.json").exists()
    assert main(["init", str(tmp_path / "e")]) == 2
    assert main(["init", "--help"]) == 0


def test_load_deck_rejects_escaping_working_copy(tmp_path, potx_file):
    d = tmp_path / "deck"
    init_deck(d, potx_file)
    victim = tmp_path / "victim.pptx"
    victim.write_text("keep")
    meta = json.loads((d / "deck.json").read_text())
    for bad in ("../victim.pptx", str(victim), "sub/x.pptx", "x.txt"):
        meta["working_copy"] = bad
        (d / "deck.json").write_text(json.dumps(meta))
        with pytest.raises(PptxSkillsError, match="working_copy"):
            load_deck(d)
        with pytest.raises(PptxSkillsError):
            changed_externally(d)
    meta["working_copy"] = "deck.pptx"
    meta["sources"] = []
    (d / "deck.json").write_text(json.dumps(meta))
    with pytest.raises(PptxSkillsError, match="sources"):
        load_deck(d)
    assert victim.read_text() == "keep"


def test_invalid_style_file_keeps_old_deck_on_replace(tmp_path, potx_file):
    d = tmp_path / "keep"
    init_deck(d, potx_file)
    before = (d / "deck.json").read_bytes(), file_hash(d / "keep.pptx")
    bad = tmp_path / "bad.json"
    bad.write_text("{nope")
    with pytest.raises(PptxSkillsError, match="not valid JSON"):
        init_deck(d, potx_file, style_file=bad, mode="replace")
    assert ((d / "deck.json").read_bytes(), file_hash(d / "keep.pptx")) == before
    assert (d / "style/brief.json").exists()


def test_invalid_style_file_leaves_no_partial_folder(tmp_path, potx_file):
    bad = tmp_path / "bad.json"
    bad.write_text("{nope")
    with pytest.raises(PptxSkillsError):
        init_deck(tmp_path / "new", potx_file, style_file=bad)
    assert not (tmp_path / "new").exists()
    assert [p.name for p in tmp_path.iterdir() if p.name.startswith(".")] == []


def test_init_refuses_to_overwrite_existing_working_copy(tmp_path, potx_file):
    d = tmp_path / "w"
    d.mkdir()
    (d / "w.pptx").write_text("precious")
    with pytest.raises(PptxSkillsError, match="--replace"):
        init_deck(d, potx_file)
    assert (d / "w.pptx").read_text() == "precious"
    init_deck(d, potx_file, mode="replace")
    assert len(open_presentation(d / "w.pptx").slides) == 0


def _with_sections(prs):
    from lxml import etree

    p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
    from conftest import layout_named

    for text in ("A", "B", "C"):
        prs.slides.add_slide(layout_named(prs, "Title Only")).shapes.title.text = text
    ids = [s.slide_id for s in prs.slides]
    rids = [e.rId for e in prs.slides._sldIdLst]
    sections = (
        f'<p14:section name="One" id="{{11111111-1111-1111-1111-111111111111}}"><p14:sldIdLst><p14:sldId id="{ids[0]}"/></p14:sldIdLst></p14:section>'
        f'<p14:section name="Two" id="{{22222222-2222-2222-2222-222222222222}}"><p14:sldIdLst><p14:sldId id="{ids[1]}"/><p14:sldId id="{ids[2]}"/></p14:sldIdLst></p14:section>'
    )
    ext = etree.fromstring(
        '<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">'
        f'<p14:sectionLst xmlns:p14="{p14}">{sections}</p14:sectionLst></p:ext></p:extLst>'
    )
    root = prs.part._element
    root.append(ext)
    show = etree.fromstring(
        '<p:custShowLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<p:custShow name="S" id="0"><p:sldLst><p:sld r:id="{rids[0]}"/><p:sld r:id="{rids[1]}"/></p:sldLst></p:custShow></p:custShowLst>'
    )
    root.find("{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst").addnext(show)


def test_delete_slide_cleans_sections_and_custom_shows(make_deck, tmp_path):
    from lxml import etree

    path = make_deck(_with_sections, name="sec.pptx")
    prs = open_presentation(path)
    gone, rid = prs.slides[0].slide_id, prs.slides._sldIdLst[0].rId
    delete_slide(prs, prs.slides[0])
    xml = etree.tostring(prs.part._element).decode()
    assert f'id="{gone}"' not in xml and f'r:id="{rid}"' not in xml
    for slide in list(prs.slides):
        delete_slide(prs, slide)
    xml = etree.tostring(prs.part._element).decode()
    assert xml.count("<p14:section ") == 2 and "<p14:sldId " not in xml and "<p:sld " not in xml
    out = tmp_path / "o.pptx"
    prs.save(str(out))
    assert len(open_presentation(out).slides) == 0


def test_init_strips_a_deck_with_a_slide_zoom(tmp_path, make_deck):
    """init removes every slide, so a zoom on one of them does not block removing its target."""
    source = linked_deck(make_deck, source=3, target=1, zoom=True, outline=1)  # the target goes first
    init_deck(tmp_path / "decks/z", source)
    working = tmp_path / "decks/z/z.pptx"
    assert len(open_presentation(working).slides) == 0
    names = zipfile.ZipFile(working).namelist()
    assert len(names) == len(set(names)) and not [n for n in names if n.startswith("ppt/slides/")]
