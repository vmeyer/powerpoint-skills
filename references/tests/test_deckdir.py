import json
import os
import zipfile
from pathlib import Path

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


def test_init_without_a_source_uses_the_neutral_template(tmp_path):
    deck = init_deck(tmp_path / "n", None)
    assert deck["sources"]["master"] is None and len(open_presentation(tmp_path / "n/n.pptx").slides) == 0


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
    assert main(["init", str(tmp_path / "e")]) == 0  # no source: the neutral template
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
    with pytest.raises(PptxSkillsError, match="new or empty folder") as exc:
        init_deck(d, potx_file)
    assert "--replace" not in str(exc.value)
    with pytest.raises(PptxSkillsError, match="not a deck folder") as exc:
        init_deck(d, potx_file, mode="replace")
    assert "--replace" not in str(exc.value)
    assert (d / "w.pptx").read_text() == "precious"


def _user_folder(tmp_path):
    """A project folder that is not a deck folder but has folders named like the tool's own."""
    proj = tmp_path / "proj"
    (proj / "style").mkdir(parents=True)
    (proj / "style/site.css").write_text("body {}")
    (proj / "check").mkdir()
    (proj / "check/notes.txt").write_text("mine")
    style = tmp_path / "s.json"
    style.write_text("{}")
    return proj, style


def test_replace_refuses_a_folder_without_deck_json(tmp_path):
    proj, style = _user_folder(tmp_path)
    for mode in ("replace", "new"):
        with pytest.raises(PptxSkillsError, match="new or empty folder") as exc:
            init_deck(proj, None, style_file=style, mode=mode)
        assert "--replace" not in str(exc.value)
    assert (proj / "style/site.css").read_text() == "body {}" and (proj / "check/notes.txt").read_text() == "mine"
    assert sorted(p.name for p in proj.iterdir()) == ["check", "style"]


def test_cli_replace_on_a_user_folder_keeps_its_files(tmp_path, capsys):
    proj, style = _user_folder(tmp_path)
    assert main(["init", str(proj), "--style-file", str(style), "--replace"]) == 2
    assert "not a deck folder" in capsys.readouterr().err
    assert (proj / "style/site.css").is_file() and (proj / "check/notes.txt").is_file()


@pytest.mark.parametrize("name", [".", "..", "My Deck", "deck_1", "Deck", "-deck", "deck-"])
def test_init_needs_a_kebab_case_folder_name(tmp_path, monkeypatch, potx_file, name):
    work = tmp_path / "work"
    work.mkdir()
    monkeypatch.chdir(work)
    with pytest.raises(PptxSkillsError, match="kebab-case"):
        init_deck(Path(name), potx_file)
    assert list(work.iterdir()) == []


def test_cli_init_dot_is_refused(tmp_path, monkeypatch, potx_file, capsys):
    monkeypatch.chdir(tmp_path)
    assert main(["init", ".", "--from", str(potx_file)]) == 2
    assert "kebab-case" in capsys.readouterr().err
    assert not (tmp_path / "deck.json").exists() and not (tmp_path / "style").exists()


def test_replace_removes_only_what_the_old_deck_describes(tmp_path, potx_file):
    d = tmp_path / "p"
    init_deck(d, potx_file)
    meta = json.loads((d / "deck.json").read_text())
    (d / "p.pptx").rename(d / "old.pptx")  # an older deck whose working copy had another name
    meta["working_copy"] = "old.pptx"
    (d / "deck.json").write_text(json.dumps(meta))
    (d / "storyline.md").write_text("plan")
    (d / "render").mkdir()
    (d / "render/slide-01.png").write_bytes(b"png")
    init_deck(d, potx_file, mode="replace")
    assert sorted(p.name for p in d.iterdir()) == ["deck.json", "p.pptx", "storyline.md", "style"]
    assert (d / "storyline.md").read_text() == "plan"


def test_replace_refuses_to_overwrite_a_file_the_old_deck_does_not_own(tmp_path, potx_file):
    d = tmp_path / "p"
    init_deck(d, potx_file)
    meta = json.loads((d / "deck.json").read_text())
    (d / "p.pptx").rename(d / "old.pptx")
    meta["working_copy"] = "old.pptx"
    (d / "deck.json").write_text(json.dumps(meta))
    (d / "p.pptx").write_text("the user's own file")
    with pytest.raises(PptxSkillsError, match="p.pptx"):
        init_deck(d, potx_file, mode="replace")
    assert (d / "p.pptx").read_text() == "the user's own file" and (d / "old.pptx").is_file()


@pytest.mark.skipif(os.geteuid() == 0, reason="root ignores folder permissions")
def test_replace_reports_files_it_cannot_remove(tmp_path, potx_file):
    d = tmp_path / "p"
    init_deck(d, potx_file)
    (d / "render").mkdir()
    (d / "render/slide-01.png").write_bytes(b"png")
    (d / "render").chmod(0o500)  # the file inside cannot be deleted
    try:
        with pytest.raises(PptxSkillsError, match="could not remove .*render"):
            init_deck(d, potx_file, mode="replace")
    finally:
        (d / "render").chmod(0o700)
    assert (d / "deck.json").is_file()  # still a deck folder, so --replace can be run again


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


def test_init_without_kept_slides_drops_the_sections(tmp_path, make_deck):
    from conftest import P14, sectioned_deck, sections

    source = sectioned_deck(make_deck)
    init_deck(tmp_path / "s", source)
    prs = open_presentation(tmp_path / "s/s.pptx")
    assert prs.part._element.find(f".//{{{P14}}}sectionLst") is None
    assert len(open_presentation(source).slides) == 3  # the source is untouched
    init_deck(tmp_path / "k", source, keep_slides=True)
    assert sections(open_presentation(tmp_path / "k/k.pptx")) == sections(open_presentation(source))
