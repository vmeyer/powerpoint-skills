"""The skills an agent reads: frontmatter, length, and every `pptx` command and link they mention."""

import json
import re
from pathlib import Path

from pptx_skills.changeset import _schema, validate_changeset
from pptx_skills.cli import main

ROOT = Path(__file__).resolve().parents[2]
SKILLS = sorted((ROOT / "skills").glob("*/SKILL.md"))
DOCS = [*SKILLS, *sorted((ROOT / "skills").glob("*/references/*.md"))]

_FENCE = re.compile(r"^```[^\n]*\n(.*?)^```", re.S | re.M)
_INLINE = re.compile(r"`([^`\n]+)`")
# `pptx <command>` as typed in a shell, not a file name such as `deck.pptx changes.json`
_COMMAND = re.compile(r"(?<![\w.$~/-])pptx ([a-z][\w-]*)")
_LINK = re.compile(r"\]\(([^)#\s]+)(?:#[^)]*)?\)")


def frontmatter(md: Path) -> dict:
    lines = md.read_text(encoding="utf-8").splitlines()
    assert lines and lines[0] == "---", f"{md} has no frontmatter"
    end = lines.index("---", 1)
    fields = {}
    for line in lines[1:end]:
        key, sep, value = line.partition(":")
        assert sep and key and not key.startswith(" "), f"{md}: unexpected frontmatter line {line!r}"
        fields[key.strip()] = value.strip()
    return fields


def code_commands(text: str) -> set[str]:
    """`pptx` subcommands used in fenced code blocks and inline code; prose such as "a pptx file" does not count."""
    code = [m.group(1) for m in _FENCE.finditer(text)]
    prose = _FENCE.sub("", text)
    code += [m.group(1) for m in _INLINE.finditer(prose)]
    return {cmd for chunk in code for cmd in _COMMAND.findall(chunk)}


def test_code_commands_only_from_code():
    text = "Open the pptx file.\n\nRun `pptx grid decks/a --cols 3`.\n\n```\npptx apply decks/a/a.pptx change.json\n```\n"
    assert code_commands(text) == {"grid", "apply"}


def test_skills_exist():
    assert "slide-design" in [md.parent.name for md in SKILLS]


def test_frontmatter_and_length():
    for md in SKILLS:
        fm = frontmatter(md)
        assert set(fm) == {"name", "description"} and fm["name"] == md.parent.name, md
        assert len(fm["description"]) <= 1024, md
        assert len(md.read_text(encoding="utf-8").splitlines()) < 300, md


def test_commands_in_skills_exist():
    for md in DOCS:
        for cmd in code_commands(md.read_text(encoding="utf-8")):
            assert main([cmd, "--help"]) == 0, f"{md.name}: pptx {cmd}"


def test_slide_design_teaches_every_command():
    folder = ROOT / "skills" / "slide-design"
    used = set().union(*(code_commands(md.read_text(encoding="utf-8")) for md in [folder / "SKILL.md", *folder.glob("references/*.md")]))
    assert used >= {"init", "style", "inventory", "apply", "grid", "render", "check", "doctor"}


def test_no_agent_specific_tool_names():
    for md in DOCS:
        assert not re.search(r"\b(Read|Write|Edit|Bash|Grep|Glob) tool\b", md.read_text(encoding="utf-8")), md


def _json_examples(text: str) -> list:
    """Objects in ```json blocks: a whole block, or one object per line."""
    found = []
    for block in re.findall(r"^```json\n(.*?)^```", text, re.S | re.M):
        try:
            found.append(json.loads(block))
        except ValueError:
            found.extend(json.loads(line) for line in block.splitlines() if line.strip())
    return found


def test_json_examples_validate():
    known = {branch["title"] for branch in _schema()["properties"]["ops"]["items"]["oneOf"]}
    checked = 0
    for md in DOCS:
        for example in _json_examples(md.read_text(encoding="utf-8")):
            changeset = example if "ops" in example else {"ops": [example]}
            ops = [o for o in changeset["ops"] if o["op"] in known]  # ops still to come are marked as such in the text
            validate_changeset({**changeset, "ops": ops} if ops else {"ops": [{"op": "add_slide", "layout": "x"}]})
            checked += len(ops)
    assert checked >= 20


def test_relative_links_resolve():
    for md in DOCS:
        for target in _LINK.findall(md.read_text(encoding="utf-8")):
            if "://" not in target:
                assert (md.parent / target).exists(), f"{md.name}: broken link {target}"
