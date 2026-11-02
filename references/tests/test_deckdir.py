import json

import pytest

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
    style = tmp_path / "s.json"
    style.write_text("{}")
    with pytest.raises(PptxSkillsError, match="not supported yet"):
        init_deck(tmp_path / "n", None, style_file=style)
    assert not (tmp_path / "n").exists()


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
