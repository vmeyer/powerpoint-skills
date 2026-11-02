"""The skills an agent reads: frontmatter, length, and every `pptx` command and link they mention."""

import json
import re
from pathlib import Path

from pptx_skills.brief import DEFAULT_RULES
from pptx_skills.changeset import validate_changeset
from pptx_skills.cli import main
from pptx_skills.neutral import validate_style

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
    assert {"slide-design", "deck-storyline"} <= {md.parent.name for md in SKILLS}


def test_frontmatter_and_length():
    for md in SKILLS:
        fm = frontmatter(md)
        assert set(fm) == {"name", "description"} and fm["name"] == md.parent.name, md
        assert len(fm["description"]) <= 1024, md
        assert len(md.read_text(encoding="utf-8").splitlines()) < 300, md
    assert len((ROOT / "skills/slide-design/SKILL.md").read_text(encoding="utf-8").splitlines()) <= 150


def test_commands_in_skills_exist():
    for md in DOCS:
        for cmd in code_commands(md.read_text(encoding="utf-8")):
            assert main([cmd, "--help"]) == 0, f"{md.name}: pptx {cmd}"


def test_slide_design_teaches_every_command():
    folder = ROOT / "skills" / "slide-design"
    used = set().union(*(code_commands(md.read_text(encoding="utf-8")) for md in [folder / "SKILL.md", *folder.glob("references/*.md")]))
    assert used >= {"init", "theme", "build", "exec", "style", "inventory", "apply", "grid", "render", "check", "doctor"}


STORYLINE_FORMAT = ROOT / "skills" / "deck-storyline" / "references" / "storyline-format.md"


def test_one_storyline_format():
    """deck-storyline defines storyline.md; slide-design points to that definition and reads existing decks the same way."""
    design = (ROOT / "skills" / "slide-design" / "SKILL.md").read_text(encoding="utf-8")
    storyline = (ROOT / "skills" / "deck-storyline" / "SKILL.md").read_text(encoding="utf-8")
    assert "../deck-storyline/references/storyline-format.md" in _LINK.findall(design)
    assert "references/storyline-format.md" in _LINK.findall(storyline)
    assert "inventory" in code_commands(storyline) and "--outline" in storyline


def test_storyline_example_follows_the_format():
    text = STORYLINE_FORMAT.read_text(encoding="utf-8")
    example = re.search(r"^```markdown\n(.*?)^```", text, re.S | re.M).group(1)
    headings = [int(n) for n in re.findall(r"^### Slide (\d+)$", example, re.M)]
    assert headings == list(range(1, len(headings) + 1)) and len(headings) >= 5
    statuses = set()
    for entry in re.split(r"^### Slide \d+\n", example, flags=re.M)[1:]:
        fields = dict(re.findall(r"^- \*\*(\w+):\*\* (.+)$", entry, re.M))
        assert {"Kicker", "Title", "Shows", "Sources", "Status", "Id"} <= set(fields), entry
        assert fields["Id"] == "—" or fields["Id"].isdigit(), entry
        assert len(fields["Title"]) <= 120, fields["Title"]  # two lines at title size, as the format says
        statuses.add(fields["Status"])
    assert statuses == {"open", "in design", "approved", "existing"}
    assert not re.search(r" [·|•] ", example)  # no keyword chains


def test_no_agent_specific_tool_names():
    for md in DOCS:
        assert not re.search(r"\b(Read|Write|Edit|Bash|Grep|Glob) tool\b", md.read_text(encoding="utf-8")), md


_BLOCK = re.compile(r"^```([^\n]*)\n(.*?)^```", re.S | re.M)


def _json_values(block: str) -> list:
    """A whole block as one JSON value, or one value per line."""
    try:
        return [json.loads(block)]
    except ValueError:
        return [json.loads(line) for line in block.splitlines() if line.strip()]


def _json_examples(text: str) -> list:
    """Objects in ```json blocks (which must parse) and in untagged blocks that are JSON (others are plain text)."""
    found = []
    for tag, block in _BLOCK.findall(text):
        if tag.strip() == "json":
            found.extend(_json_values(block))
        elif not tag.strip():
            try:
                values = _json_values(block)
            except ValueError:
                continue  # grid output and the like
            found.extend(v for v in values if isinstance(v, dict))
    return found


def _changesets(text: str) -> list[dict]:
    """Every example change set; a lone op counts as a change set of one op."""
    return [e if "ops" in e else {"ops": [e]} for e in _json_examples(text) if "op" in e or "ops" in e]


def validate_examples(text: str) -> int:
    """Validate every example (change sets against the schema, the rest as style files); returns the ops checked."""
    checked = 0
    for example in _json_examples(text):
        if "op" not in example and "ops" not in example:
            validate_style(example)
            continue
        changeset = example if "ops" in example else {"ops": [example]}
        validate_changeset(changeset)
        checked += len(changeset["ops"])
    return checked


def test_json_examples_validate():
    assert sum(validate_examples(md.read_text(encoding="utf-8")) for md in DOCS) >= 20


def test_examples_with_unknown_ops_fail_also_in_untagged_blocks():
    import pytest

    from pptx_skills.errors import PptxSkillsError

    for fence in ("```json", "```"):
        with pytest.raises(PptxSkillsError, match="unknown op"):
            validate_examples(f'{fence}\n{{"op": "add_slide", "layout": "Blank"}}\n{{"op": "distribute", "slide": 1}}\n```\n')
    assert validate_examples("```\n1.69,4.45 9.96x6.14 | 11.95,4.45 9.96x6.14\n```\n") == 0


_CREATES = ("add_slide", "clone_slide", "replace_slide")


def test_examples_use_action_titles_and_create_last_first():
    """Example titles are statements, and a move of "$last" shows the op that creates that slide."""
    min_words = DEFAULT_RULES["title_min_words"]
    for md in DOCS:
        for changeset in _changesets(md.read_text(encoding="utf-8")):
            ops = changeset["ops"]
            for n, spec in enumerate(ops):
                if spec["op"] == "set_text" and spec["target"] == {"role": "title"}:
                    text = spec.get("text") or " ".join(r.get("t", "") for p in spec["paragraphs"] for r in p["runs"])
                    assert len(text.split()) >= min_words, f"{md.name}: topic title {text!r}"
                if spec["op"] == "move_slide" and spec["slide"] == "$last":
                    assert any(o["op"] in _CREATES for o in ops[:n]), f"{md.name}: move of $last without the op that creates it"


def test_relative_links_resolve():
    for md in DOCS:
        for target in _LINK.findall(md.read_text(encoding="utf-8")):
            if "://" not in target:
                assert (md.parent / target).exists(), f"{md.name}: broken link {target}"


def _body_sizes(changeset: dict) -> set:
    """Font sizes a change set gives text other than kicker and title."""
    sizes = set()
    for spec in changeset["ops"]:
        if spec.get("target") in ({"role": "title"}, {"role": "kicker"}):
            continue
        sizes |= {(spec.get("font") or {}).get("size")}
        sizes |= {run.get("size") for p in spec.get("paragraphs", []) for run in p["runs"]}
    return sizes - {None}


def test_slide_skeletons_use_at_most_two_body_sizes():
    """Design rule 10: title plus one or two body sizes; a key figure is the larger of the two."""
    for md in DOCS:
        for changeset in _changesets(md.read_text(encoding="utf-8")):
            assert len(_body_sizes(changeset)) <= 2, f"{md.name}: {changeset.get('label')} uses sizes {sorted(_body_sizes(changeset))}"


def test_setup_docs_cover_uv_the_plugin_folder_and_the_first_run():
    """An agent can set the engine up from the README or the skill alone, also for an installed plugin."""
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    skill = (ROOT / "skills" / "slide-design" / "SKILL.md").read_text(encoding="utf-8")
    for name, text in (("README.md", readme), ("SKILL.md", skill)):
        for needle in ("brew install uv", "astral.sh/uv/install.sh", "plugins/cache", "network"):
            assert needle in text, f"{name} lacks {needle!r}"
    for needle in ("pptx init decks/<slug>`", "pptx theme", "pptx build", "pptx exec", "--strict", "--sheets"):
        assert needle in readme, f"README.md lacks {needle!r}"


def _python_examples(text: str) -> list[str]:
    return [block for tag, block in _BLOCK.findall(text) if tag.strip() == "python"]


def test_script_examples_run(tmp_path):
    """Every python example in scripts.md runs through `pptx exec` on a neutral deck, in order, without errors."""
    from pptx_skills.checks import run_checks
    from pptx_skills.deckdir import init_deck
    from pptx_skills.scripting import run_script

    blocks = _python_examples((ROOT / "skills/slide-design/references/scripts.md").read_text(encoding="utf-8"))
    assert len(blocks) >= 2
    init_deck(tmp_path / "decks/s", None)
    deck = tmp_path / "decks/s/s.pptx"
    for n, block in enumerate(blocks, 1):
        script = tmp_path / f"example-{n}.py"
        script.write_text(block, encoding="utf-8")
        assert run_script(deck, script)["touched"], f"example {n} changed nothing"
    report = run_checks(deck, all_slides=True)
    assert report["summary"]["error"] == 0, report["findings"]


def test_self_contained_examples_build_without_errors(tmp_path):
    """Example change sets that create their own slide apply to a neutral deck and pass the default check."""
    from pptx_skills.changeset import apply_changeset
    from pptx_skills.checks import run_checks
    from pptx_skills.deckdir import init_deck

    init_deck(tmp_path / "decks/x", None)
    deck = tmp_path / "decks/x/x.pptx"
    built = 0
    for md in DOCS:
        for changeset in _changesets(md.read_text(encoding="utf-8")):
            ops = changeset["ops"]
            if ops[0]["op"] != "add_slide" or any(o.get("slide", "$last") != "$last" for o in ops):
                continue
            ops = [dict(o, layout="Title Only") if o["op"] == "add_slide" else o for o in ops]  # the neutral layout names
            apply_changeset(deck, {"ops": ops})
            report = run_checks(deck)
            assert report["summary"]["error"] == 0 and report["summary"]["warning"] == 0, (md.name, changeset.get("label"), report["findings"])
            built += 1
    assert built >= 4
