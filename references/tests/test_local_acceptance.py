"""Acceptance tests against private files. They run only where the env vars point to such files."""

import json
import os
from pathlib import Path

import pytest

from pptx_skills import cli
from pptx_skills.brief import load_brief
from pptx_skills.changeset import apply_changeset
from pptx_skills.checks import run_checks
from pptx_skills.deckdir import init_deck
from pptx_skills.errors import PptxSkillsError
from pptx_skills.opc import open_presentation
from pptx_skills.render import find_pdftoppm, find_soffice, render

MASTER, EXAMPLES, REFERENCE = ("PPTX_SKILLS_LOCAL_MASTER", "PPTX_SKILLS_LOCAL_EXAMPLES", "PPTX_SKILLS_LOCAL_REFERENCE")


def needs(*env_vars):
    missing = [v for v in env_vars if not (os.environ.get(v) and Path(os.environ[v]).is_file())]
    return pytest.mark.skipif(bool(missing), reason=f"set {', '.join(missing)} to a local file to run this test")


def _path(var):
    return Path(os.environ[var])


@needs(REFERENCE)
def test_reference_slide_passes():
    rep = run_checks(_path(REFERENCE), all_slides=True)
    assert rep["summary"]["error"] == 0, [f for f in rep["findings"] if f["severity"] == "error"]
    assert not [f for f in rep["findings"] if f["category"] == "text" and f["severity"] == "warning"]


@needs(REFERENCE)
@pytest.mark.real_render
@pytest.mark.skipif(not (find_soffice() and find_pdftoppm()), reason="LibreOffice or poppler not installed")
def test_reference_slide_renders(tmp_path):
    files = render(_path(REFERENCE), tmp_path / "out")
    assert files and all(f.is_file() and f.stat().st_size > 0 for f in files)


@needs(MASTER)
def test_master_brief(tmp_path):
    init_deck(tmp_path / "m", _path(MASTER))
    brief = load_brief(tmp_path / "m")
    assert brief["content_area_source"] in ("guides", "layouts")
    assert any(p["role"] == "kicker" for lay in brief["layouts"] for p in lay["placeholders"])


@needs(MASTER, EXAMPLES)
def test_clone_every_example(tmp_path):
    deck_dir = tmp_path / "c"
    init_deck(deck_dir, _path(MASTER), _path(EXAMPLES))
    working = deck_dir / "c.pptx"
    total = len(open_presentation(_path(EXAMPLES)).slides)
    assert total > 0
    cloned, refused = 0, []
    for number in range(1, total + 1):
        op = {"op": "clone_slide", "from": {"deck": "examples", "slide": number}}
        try:
            apply_changeset(working, {"ops": [op]})
        except PptxSkillsError as exc:
            refused.append((number, str(exc)))
        else:
            cloned += 1
    assert cloned > 0, f"cloned {cloned} of {total} example slides, refused {len(refused)}: {refused}"
    assert all("cannot be cloned" in message for _, message in refused), refused
    assert len(open_presentation(working).slides) == cloned
    report = run_checks(working, all_slides=True)
    assert [f for f in report["findings"] if f["category"] == "structure"] == []


def _title_only_layout(brief):
    """A layout whose only content placeholders are kicker and title (footer, date and number are copied)."""
    plain = {"kicker", "title", "number", "footer", "date"}
    for lay in brief["layouts"]:
        roles = {p["role"] for p in lay["placeholders"]}
        if {"kicker", "title"} <= roles <= plain:
            return lay["name"]
    pytest.fail("no layout with kicker and title and nothing else to fill")


@needs(MASTER, EXAMPLES)
@pytest.mark.real_render
def test_cli_workflow_end_to_end(tmp_path, capsys):
    deck_dir = tmp_path / "wf"
    assert cli.main(["init", str(deck_dir), "--from", str(_path(MASTER)), "--examples", str(_path(EXAMPLES))]) == 0
    brief = load_brief(deck_dir)
    layout = _title_only_layout(brief)
    x, y, w, h = brief["content_area"]
    gap = 0.5
    box_w, box_h, box_y = round((w - gap) / 2, 2), round(h * 0.4, 2), round(y + h * 0.3, 2)
    boxes = [[round(x, 2), box_y, box_w, box_h], [round(x + box_w + gap, 2), box_y, box_w, box_h]]
    ops = [
        {"op": "add_slide", "layout": layout},
        {"op": "set_text", "slide": "$last", "target": {"role": "kicker"}, "text": "Acceptance"},
        {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "The workflow runs end to end on this master"},
        {"op": "add_shape", "slide": "$last", "name": "box-a", "shape": "rect", "box": boxes[0],
         "fill": "lt2", "line": "none", "text": "First", "font": {"size": 12, "color": "tx1"}},
        {"op": "add_shape", "slide": "$last", "name": "box-b", "shape": "rect", "box": boxes[1],
         "fill": "lt2", "line": "none", "text": "Second", "font": {"size": 12, "color": "tx1"}},
    ]
    changes = tmp_path / "change.json"
    changes.write_text(json.dumps({"label": "acceptance", "ops": ops}), encoding="utf-8")
    working = deck_dir / "wf.pptx"
    assert cli.main(["apply", str(working), str(changes)]) == 0
    capsys.readouterr()
    report = run_checks(deck_dir, all_slides=True)
    assert report["summary"]["error"] == 0, [f for f in report["findings"] if f["severity"] == "error"]
    assert not [f for f in report["findings"] if f["category"] == "text" and f["severity"] == "warning"], [f for f in report["findings"] if f["category"] == "text"]
    assert cli.main(["check", str(deck_dir)]) == 0
    assert len(open_presentation(working).slides) == 1
    if find_soffice() and find_pdftoppm():
        assert cli.main(["render", str(working), "--out", str(tmp_path / "png")]) == 0
        assert list((tmp_path / "png").glob("*.png"))
