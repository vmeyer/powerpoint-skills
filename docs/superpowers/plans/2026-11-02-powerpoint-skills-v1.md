# powerpoint-skills v1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Two skills (`deck-storyline`, `slide-design`) plus a `pptx` CLI that lets any coding agent design PowerPoint slides one at a time on a working copy, in the style of any master, with a render-check-fix loop.

**Architecture:** A Python package `pptx_skills` under `references/` exposes one CLI (`pptx`) with subcommands `init`, `style`, `inventory`, `apply`, `grid`, `render`, `check`, `doctor`. The `.pptx` working copy is the source of truth; agents change it only through validated JSON change sets applied atomically by `apply`. Rendering is LibreOffice → PDF → PNG via poppler; checks are deterministic rules on the XML plus font-metric text measurement.

**Tech Stack:** Python ≥ 3.11, uv, python-pptx 1.0.2, lxml, Pillow ≥ 10.1, jsonschema, pytest, hatchling (build backend); external processes: LibreOffice (`soffice`), poppler (`pdftoppm`).

**Spec:** `docs/superpowers/specs/2026-11-02-powerpoint-skills-v1-design.md` (read it first; this plan argues from it).

## Global Constraints

- Never commit `.pptx`, `.potx`, `.pptm`, `.ppt`, `.odp`, `decks/` or screenshots; the repo is public. Repo text names no company; the local test case is "a corporate master".
- Clean room: nothing from Anthropic's `pptx` skill is opened, copied or paraphrased. Sources: ECMA-376, python-pptx, LibreOffice and poppler docs.
- Runtime is local and offline. Dependencies: python-pptx (MIT), lxml, Pillow, jsonschema only. No PyMuPDF (AGPL).
- Every CLI command: `--help`; exit codes 0 = ok, 1 = check findings, 2 = usage or runtime error; `--json` for machine-readable output. Errors go to stderr as `error: <message>`.
- Units in change sets and output: centimetres, rounded to 2 decimals. Colors: theme roles (`dk1 lt1 dk2 lt2 tx1 tx2 bg1 bg2 accent1…accent6 hlink folHlink`) with optional brightness `+NN` (lighter) or `-NN` (darker); raw `#RRGGBB` only with `"allow_rgb": true`.
- Check defaults (exact): `min_size_pt` 11, `min_size_small_pt` 9, `max_lines` 3 (per paragraph), `max_title_lines` 2, `max_bullets` 5, `near_align_cm` 0.3, `snap_cm` 0.05, `title_min_words` 4.
- Output convention in the user's project: `decks/<slug>/…` as in spec §5.3. `init` never overwrites an existing deck folder silently.
- SKILL.md files: frontmatter only `name` (= folder) and `description`; under 300 lines; only shell commands, no agent-specific tool names.
- Commits: Conventional Commits in English, identity `Volker Meyer <5428860+vmeyer@users.noreply.github.com>` (check `git config user.email` once), work on branch `feat/slide-design`.
- Test command (from repo root): `(cd references && uv run pytest tests/<file> -v)`.
- Fixtures and helpers named in test snippets (`deck`, `deck_with_two_slides`, `apply(deck, ops)`, `inv_shape(name)`, `shape_named(name)`, `finding(report, rule)`, `exists(name)`, `layout_named(prs, name)`) are added to `references/tests/conftest.py` by the task that first uses them; they are thin wrappers over the interfaces below and generate every file at test time.

## Review Focus

1. **Shapes inside groups.** Example decks are full of groups (one corporate example deck has 540). Inventory boxes, targets and every geometry check must use absolute slide coordinates, including nested groups with `chOff`/`chExt` scaling. Pinned by `test_group_children_absolute_boxes` (Task 3) and `test_off_slide_child_in_group` (Task 14).
2. **Inherited text formatting.** Most runs carry no explicit size or font; overflow and minimum-size checks must use the effective style from layout, master and theme, never treat a missing size as 0 or 18. Pinned by `test_effective_style_inherits_from_master` (Task 10).
3. **Ordinary separators that are not keyword chains.** A footer such as `|   Company AG`, "z. B.", "Pipelines, Gates, Audit-Trail" must not be flagged; only three or more items joined by ` · `, ` | ` or ` • ` are. Pinned by `test_keyword_chain_only_for_three_items` (Task 11).
4. **Working copy changed or open in PowerPoint.** A hand-edited working copy (hash differs) or a lock file next to it must not break `apply`; it applies to the current file state and reports the condition. The source deck given to `init` is never written. Pinned by `test_apply_after_external_edit_reports_and_applies` (Task 6) and `test_init_never_writes_source` (Task 5).
5. **Theme colors via `tx1`/`bg1` aliases and master-inherited RGB.** `schemeClr` of any role counts as a theme color; only `srgbClr` written on the slide itself is reported. Pinned by `test_rgb_rule_ignores_scheme_aliases` (Task 14).

## File Structure

```
references/
  pyproject.toml, uv.lock
  pptx_skills/
    __init__.py
    cli.py            # argparse, subcommand registry, exit codes
    errors.py         # PptxSkillsError
    units.py          # cm/EMU/pt/guide conversions
    geometry.py       # Box, absolute boxes, snapping
    opc.py            # open .pptx/.potx, save, hash, lock file
    colors.py         # color reference grammar, read/apply
    roles.py          # placeholder role inference
    inventory.py      # inventory JSON + markdown, `inventory` command
    brief.py          # style brief, `style` command (data part)
    deckdir.py        # deck folder, deck.json, snapshots, change log, `init` command
    slides.py         # slide refs, add/delete/move, clone/replace across decks
    changeset.py      # schema validation, op registry, atomic apply, `apply` command
    ops_text.py       # set_text
    ops_shapes.py     # add_shape, add_line, update_shape, delete_shape, z_order, group, ungroup
    ops_tables.py     # add_table, set_table
    ops_media.py      # add_chart, add_picture, set_notes
    grid.py           # `grid` command
    fonts.py          # font lookup, effective run style, text measurement
    checks/
      __init__.py     # run_checks, report, `check` command
      structure.py, text.py, geometry.py, style.py
    render.py         # LibreOffice/poppler, contact sheets, `render` and `doctor` commands
    neutral.py        # neutral template from a style file
    schemas/changeset.schema.json, style-file.schema.json
  tests/
    conftest.py       # generated fixtures; local-only fixtures via env vars
    test_*.py
skills/slide-design/SKILL.md + references/{changeset-format,visual-forms,review-checklist}.md
skills/deck-storyline/SKILL.md + references/storyline-format.md
scripts/sync-mirrors.sh, .codex-plugin/plugin.json (generated)
```

Schemas live inside the package (not `references/schemas/` as sketched in spec §5.4) so they resolve the same way however the package is installed.

**Milestone M1** is reached after Task 13: Volker can use `slide-design` on the corporate master with text and shapes. Tasks 14–21 complete v1.

---

### Task 1: Package, CLI skeleton and file access

**Files:**
- Create: `references/pyproject.toml`, `references/pptx_skills/{__init__,cli,errors,units,geometry,opc}.py`, `references/tests/conftest.py`
- Test: `references/tests/test_foundation.py`

**Interfaces:**
- Produces:
  - `errors.PptxSkillsError(Exception)`
  - `cli.main(argv: list[str] | None = None) -> int`; `cli.OK, cli.FINDINGS, cli.ERROR = 0, 1, 2`; `cli.COMMAND_MODULES: list[ModuleType]` — each module has `register(subparsers) -> None` that adds its subparser with `set_defaults(handler=fn)`, `fn(args) -> int`; `cli.run_handler(handler, args) -> int` maps `PptxSkillsError` to stderr `error: …` and exit 2; `cli.emit(args, data: dict, text: str) -> None` prints JSON when `args.json`, else `text`.
  - `units.EMU_PER_CM = 360000`, `EMU_PER_PT = 12700`; `cm_to_emu(cm: float) -> int`; `emu_to_cm(emu: int) -> float` (2 decimals); `guide_to_cm(pos: int) -> float` (guide pos is 1/8 pt).
  - `geometry.Box(x, y, w, h)` frozen dataclass in cm with `right`, `bottom`, `contains(other, tol=0.01)`, `intersects(other) -> float` (overlap area), `as_list() -> list[float]`.
  - `opc.open_presentation(path: Path) -> Presentation`; `opc.save_presentation(prs, path: Path) -> str` (sha256 of written file); `opc.file_hash(path) -> str`; `opc.lock_file(path) -> Path | None`.
  - conftest fixtures: `default_pptx` (python-pptx default saved to tmp), `potx_file` (same, content type rewritten to template), `make_deck(builder)` factory.

- [ ] **Step 1: Write `pyproject.toml`**: project `pptx-skills` 0.1.0, `requires-python = ">=3.11"`, deps `python-pptx>=1.0.2`, `lxml>=5`, `Pillow>=10.1`, `jsonschema>=4.21`; `[project.scripts] pptx = "pptx_skills.cli:main"`; hatchling build with `packages = ["pptx_skills"]` and the `schemas/` folder included; dev group `pytest>=8`. Run `(cd references && uv sync)`; expected: creates `.venv`, `uv.lock`.

- [ ] **Step 2: Write the failing tests**

```python
def test_units():
    assert cm_to_emu(1.5) == 540000
    assert emu_to_cm(12192000) == 33.87
    assert guide_to_cm(339) == pytest.approx(1.5, abs=0.01)

def test_box_contains_and_intersects():
    a, b = Box(0, 0, 10, 10), Box(2, 2, 3, 3)
    assert a.contains(b) and not b.contains(a)
    assert Box(0, 0, 2, 2).intersects(Box(1, 1, 2, 2)) == pytest.approx(1.0)

def test_open_potx(potx_file):
    prs = open_presentation(potx_file)
    assert len(prs.slides) == 0 and "Title Only" in [l.name for l in prs.slide_layouts]

def test_open_rejects_other_files(tmp_path):
    (tmp_path / "x.docx").write_bytes(b"")
    with pytest.raises(PptxSkillsError, match="not a .pptx or .potx"):
        open_presentation(tmp_path / "x.docx")

def test_save_returns_hash(tmp_path, default_pptx):
    digest = save_presentation(open_presentation(default_pptx), tmp_path / "o.pptx")
    assert digest == file_hash(tmp_path / "o.pptx") and len(digest) == 64

def test_lock_file(default_pptx):
    (default_pptx.parent / f"~${default_pptx.name}").write_bytes(b"x")
    assert lock_file(default_pptx).name == f"~${default_pptx.name}"

def test_cli_exit_codes(capsys):
    assert main(["no-such-command"]) == 2
    assert main(["--help"]) == 0
    def boom(args): raise PptxSkillsError("bad input")
    assert run_handler(boom, argparse.Namespace()) == 2
    assert "error: bad input" in capsys.readouterr().err
```

- [ ] **Step 3: Run** `(cd references && uv run pytest tests/test_foundation.py -v)` — expected: FAIL (imports missing).

- [ ] **Step 4: Implement** the modules. `open_presentation` for `.potx`: read the zip into memory, replace the `/ppt/presentation.xml` override content type `application/vnd.openxmlformats-officedocument.presentationml.template.main+xml` with `…presentation.main+xml`, open from `BytesIO`. Any other extension → `PptxSkillsError("<path> is not a .pptx or .potx file")`. `main` catches argparse's `SystemExit` and returns its code (0 for `--help`, 2 for usage errors).

- [ ] **Step 5: Run** the tests — expected: PASS.

- [ ] **Step 6: Commit** `git add references && git commit -m "feat(pptx): add package skeleton, CLI and file access"`

---

### Task 2: Color references

**Files:**
- Create: `references/pptx_skills/colors.py`
- Test: `references/tests/test_colors.py`

**Interfaces:**
- Consumes: `PptxSkillsError`.
- Produces:
  - `ColorRef(role: str | None, rgb: str | None, brightness: int)` frozen dataclass; `brightness` in −99…99.
  - `parse_color(text: str, allow_rgb: bool = False) -> ColorRef | None` (`"none"` → `None`).
  - `format_color(ref: ColorRef) -> str` (inverse of parse).
  - `apply_color(color_format: pptx.dml.color.ColorFormat, ref: ColorRef) -> None`.
  - `read_color(fill_el) -> str | None`: `fill_el` is an element holding `a:schemeClr`, `a:srgbClr` or `a:sysClr` as child (e.g. `a:solidFill`); returns `"accent1+60"`, `"#1F4E79"`, or `None`.
  - `ROLE_TO_MSO: dict[str, MSO_THEME_COLOR]` (`dk1→DARK_1`, `tx1→TEXT_1`, `bg1→BACKGROUND_1`, `hlink→HYPERLINK`, `folHlink→FOLLOWED_HYPERLINK`, …).

- [ ] **Step 1: Write the failing tests**

```python
@pytest.mark.parametrize("text,ref", [
    ("accent1", ColorRef("accent1", None, 0)), ("accent1+60", ColorRef("accent1", None, 60)),
    ("dk2-25", ColorRef("dk2", None, -25)), ("tx1", ColorRef("tx1", None, 0))])
def test_parse(text, ref):
    assert parse_color(text) == ref and format_color(ref) == text

def test_none_and_errors():
    assert parse_color("none") is None
    with pytest.raises(PptxSkillsError, match="allow_rgb"): parse_color("#FF0000")
    with pytest.raises(PptxSkillsError, match="unknown color"): parse_color("accent9")
    assert parse_color("#ff0000", allow_rgb=True) == ColorRef(None, "FF0000", 0)

def test_apply_then_read(default_pptx):
    prs = Presentation(default_pptx); s = prs.slides.add_slide(prs.slide_layouts[6])
    shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, 100, 100); shp.fill.solid()
    apply_color(shp.fill.fore_color, parse_color("accent1+60"))
    assert read_color(shp._element.spPr.find(qn("a:solidFill"))) == "accent1+60"

def test_read_lummod_only_and_sysclr():
    el = parse_xml(f'<a:solidFill {nsdecls("a")}><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>')
    assert read_color(el) == "accent1-25"
    el = parse_xml(f'<a:solidFill {nsdecls("a")}><a:sysClr val="windowText" lastClr="000000"/></a:solidFill>')
    assert read_color(el) == "#000000"
```

- [ ] **Step 2: Run** `(cd references && uv run pytest tests/test_colors.py -v)` — expected: FAIL.
- [ ] **Step 3: Implement.** Brightness maps to python-pptx `ColorFormat.brightness = n/100`. Reading: `lumOff` present → `+lumOff/1000`; else `lumMod < 100000` → `-(100 - lumMod/1000)`; round to int.
- [ ] **Step 4: Run** — expected: PASS.
- [ ] **Step 5: Commit** `git commit -m "feat(pptx): parse, apply and read theme color references"`

---

### Task 3: Placeholder roles and inventory

**Files:**
- Create: `references/pptx_skills/roles.py`, `references/pptx_skills/inventory.py`; Modify: `geometry.py` (`absolute_box`), `cli.py` (register)
- Test: `references/tests/test_inventory.py`; conftest fixtures `deck_with_two_slides`, `kicker_master` (default template whose "Title Only" layout gets an extra body placeholder 0.6 cm high whose bottom edge is above the title's top edge; title moved down to make room)

**Interfaces:**
- Consumes: `Box`, `emu_to_cm`, `read_color`, `open_presentation`.
- Produces:
  - `roles.placeholder_role(ph_type, box: Box, title_box: Box | None) -> str` — one of `title subtitle kicker body picture table chart footer number date other`. Kicker = a body/object placeholder with height ≤ 1.5 cm whose bottom ≤ the title's top.
  - `geometry.absolute_box(shape) -> Box` — walks parent groups, applying each `grpSpPr/xfrm` (`off`, `ext`, `chOff`, `chExt`) scaling.
  - `inventory.inventory(prs, slides: list[int] | None = None) -> dict` with shape `{"slide_size_cm": [w, h], "slides": [{"number", "id", "layout", "title", "kicker", "notes", "shapes": [...]}]}`; shape dict keys: `id, name, kind (placeholder|autoshape|textbox|picture|table|chart|group|connector|other), geometry (prst or null), role, box [x,y,w,h], rotation, fill, line, text, alt, children` (`text` = `{"paragraphs": [{"level": int, "runs": [{"t": str, "b"?, "i"?, "color"?, "size"?}]}], "autofit": "none"|"shape"|"shrink"|null}`; only explicit run attributes are listed; `fill` is `"none"`, a color string, `"pattern:<prst> <fg> <bg>"` or null when inherited).
  - `inventory.to_markdown(inv: dict) -> str` — one heading per slide `## Slide <n> (id <id>) — <layout>`, one line per shape `- [<id>] <name> <kind>/<role> @ x,y w×h: "<text…>"`.
  - CLI: `pptx inventory <deck> [--slides 1,3] [--json]`.

- [ ] **Step 1: Write the failing tests**

```python
def test_title_role_and_ids(default_pptx):
    prs = Presentation(default_pptx); s = prs.slides.add_slide(prs.slide_layouts[5]); s.shapes.title.text = "Hello world"
    sl = inventory(prs)["slides"][0]
    assert (sl["number"], sl["id"], sl["layout"], sl["title"]) == (1, 256, "Title Only", "Hello world")
    title = next(x for x in sl["shapes"] if x["role"] == "title")
    assert title["box"] == [1.27, 0.76, 22.86, 3.17]

def test_runs_and_fill(default_pptx): ...
    # rectangle filled accent1+60 with paragraph runs "Bold"(bold) + " rest"
    assert shape["fill"] == "accent1+60"
    assert shape["text"]["paragraphs"][0]["runs"] == [{"t": "Bold", "b": True}, {"t": " rest"}]

def test_group_children_absolute_boxes(default_pptx):
    # group xfrm off (5cm,5cm) ext (4cm,2cm) chOff (0,0) chExt (2cm,1cm); child off (0,0) ext (2cm,1cm)
    assert group["children"][0]["box"] == [5.0, 5.0, 4.0, 2.0]

def test_kicker_role(kicker_master):
    prs = Presentation(kicker_master); s = prs.slides.add_slide(layout_named(prs, "Title Only"))
    assert any(x["role"] == "kicker" for x in inventory(prs)["slides"][0]["shapes"])

def test_filter_markdown_and_cli(deck_with_two_slides, capsys):
    assert [s["number"] for s in inventory(Presentation(deck_with_two_slides), slides=[2])["slides"]] == [2]
    assert main(["inventory", str(deck_with_two_slides), "--json"]) == 0
    assert json.loads(capsys.readouterr().out)["slides"][1]["number"] == 2
    assert "## Slide 1 (id 256)" in to_markdown(inventory(Presentation(deck_with_two_slides)))
```

- [ ] **Step 2: Run** — expected: FAIL.
- [ ] **Step 3: Implement** `roles.py`, `absolute_box`, `inventory.py`, register the command. Inventory reads the slide's own XML only (no inheritance), except `box` which falls back to the layout placeholder's position when the slide placeholder has no `xfrm`.
- [ ] **Step 4: Run** — expected: PASS.
- [ ] **Step 5: Commit** `git commit -m "feat(pptx): inventory slides and shapes with roles and absolute boxes"`

---

### Task 4: Style brief

**Files:**
- Create: `references/pptx_skills/brief.py`; Modify: `cli.py`
- Test: `references/tests/test_brief.py`; conftest fixture `guided_master` (default template with master `p15:sldGuideLst` guides: vertical 339 and 7340, horizontal 1094 and 3789)

**Interfaces:**
- Consumes: `open_presentation`, `placeholder_role`, `read_color`, `inventory`, `guide_to_cm`.
- Produces:
  - `DEFAULT_RULES: dict` with the exact values from Global Constraints.
  - `build_brief(master: Presentation, examples: Presentation | None = None, rules_override: dict | None = None) -> dict` with keys `slide_size_cm`, `theme {colors {role: "#RRGGBB"}, fonts {major, minor}}`, `text_styles {title_pt, body_pt: [per level]}`, `guides {vertical: [cm], horizontal: [cm]}`, `content_area [x,y,w,h]`, `content_area_source ("guides"|"layouts")`, `layouts [{name, placeholders: [{idx, type, role, box, prompt}]}]`, `examples [{number, layout, shape_count, fills {ref: n}, text_colors {ref: n}, sizes_pt [..], table_styles [..]}]`, `color_usage {fills {ref: n}, text {ref: n}}`, `typical_gap_cm {horizontal, vertical}`, `corner_radii [..]`, `rules`.
  - `write_brief(brief: dict, style_dir: Path) -> None` — writes `brief.json` and `brief.md`; keeps the text between `<!-- agent:building-blocks -->` and `<!-- /agent:building-blocks -->` from an existing `brief.md`.
  - `load_brief(deck_dir: Path) -> dict | None`.
  - CLI: `pptx style <deck-dir> [--refresh]` (data part; contact sheets are added in Task 12).

- [ ] **Step 1: Write the failing tests**

```python
def test_theme_and_size(default_pptx):
    b = build_brief(Presentation(default_pptx))
    assert b["slide_size_cm"] == [25.4, 19.05]
    assert b["theme"]["fonts"] == {"major": "Calibri", "minor": "Calibri"}
    assert b["theme"]["colors"]["accent1"] == "#4F81BD" and b["theme"]["colors"]["dk1"] == "#000000"
    assert b["text_styles"]["title_pt"] == 44 and b["text_styles"]["body_pt"][:2] == [32, 28]

def test_content_area_from_guides(guided_master):
    b = build_brief(open_presentation(guided_master))
    assert b["content_area_source"] == "guides"
    assert b["content_area"] == pytest.approx([1.5, 4.82, 30.87, 11.88], abs=0.02)

def test_content_area_from_layouts(default_pptx):
    b = build_brief(Presentation(default_pptx))
    assert (b["content_area_source"], b["content_area"]) == ("layouts", [1.27, 4.45, 22.86, 12.57])

def test_rules_default_and_override(default_pptx):
    r = build_brief(Presentation(default_pptx), rules_override={"min_size_pt": 10})["rules"]
    assert r == {**DEFAULT_RULES, "min_size_pt": 10}

def test_color_usage_from_examples(default_pptx, examples_with_three_accent1_rects):
    b = build_brief(Presentation(default_pptx), Presentation(examples_with_three_accent1_rects))
    assert b["color_usage"]["fills"]["accent1"] == 3

def test_agent_section_survives_rewrite(tmp_path, default_pptx):
    b = build_brief(Presentation(default_pptx)); write_brief(b, tmp_path)
    md = (tmp_path / "brief.md").read_text().replace("<!-- /agent:building-blocks -->", "example 3: KPI tiles\n<!-- /agent:building-blocks -->")
    (tmp_path / "brief.md").write_text(md); write_brief(b, tmp_path)
    assert "example 3: KPI tiles" in (tmp_path / "brief.md").read_text()
```

- [ ] **Step 2: Run** — expected: FAIL.
- [ ] **Step 3: Implement.** Content area from guides only when the master (or presentation) has exactly two vertical and two horizontal guides; otherwise from the layout whose placeholders are exactly one title plus one body/object (largest such body wins): area = that body's box. `sysClr` colors use `lastClr`. Typical gap: median of gaps 0 < g < 3 cm between horizontally adjacent shapes whose vertical spans overlap by ≥ 50 % (and the transposed rule for vertical), over all example slides.
- [ ] **Step 4: Run** — expected: PASS.
- [ ] **Step 5: Commit** `git commit -m "feat(pptx): build the style brief from masters and example decks"`

---

### Task 5: Deck folder and `init`

**Files:**
- Create: `references/pptx_skills/deckdir.py`, `references/pptx_skills/slides.py` (only `delete_slide` for now); Modify: `cli.py`
- Test: `references/tests/test_deckdir.py`

**Interfaces:**
- Consumes: `open_presentation`, `save_presentation`, `file_hash`, `build_brief`, `write_brief`.
- Produces `slides.delete_slide(prs, slide) -> None`: removes the `sldId` and drops its relationship (Task 6 extends `slides.py`).
- Produces:
  - `init_deck(deck_dir: Path, source: Path | None, examples: Path | None = None, style_file: Path | None = None, keep_slides: bool = False, mode: str = "new") -> dict` (returns deck.json content). `mode`: `new` (refuse if folder has a `deck.json`), `reuse` (keep working copy, refresh brief), `replace` (start over). Source kinds: `.potx` → layouts only; `.pptx` + `keep_slides` → all slides kept; `.pptx` without → slides removed; no source + `style_file` → neutral template (until Task 18 this raises `PptxSkillsError("decks from a style file alone are not supported yet")`; Task 18 replaces it).
  - deck.json keys: `slug, working_copy, sources {master, examples, style_file}` (absolute paths as given), `keep_slides, last_hash, seq, last_touched` (slide ids).
  - `load_deck(deck_dir) -> dict`; `find_deck_dir(pptx_path: Path) -> Path | None` (parent with `deck.json`).
  - `record_apply(deck_dir, changeset: dict, label: str, before: bytes, touched: list[int], new_hash: str) -> int` — writes `changes/NNN-<label>.json`, `snapshots/NNN.pptx`, updates `seq`, `last_hash`, `last_touched`; returns NNN.
  - `changed_externally(deck_dir) -> bool`.
  - CLI: `pptx init <deck-dir> [--from P] [--examples P] [--style-file P] [--keep-slides] [--reuse | --replace]`.

- [ ] **Step 1: Write the failing tests**

```python
def test_init_from_potx(tmp_path, potx_file):
    d = init_deck(tmp_path / "decks/ai-gov", potx_file)
    prs = open_presentation(tmp_path / "decks/ai-gov/ai-gov.pptx")
    assert len(prs.slides) == 0 and d["slug"] == "ai-gov"
    assert (tmp_path / "decks/ai-gov/style/brief.json").exists()

def test_init_strips_or_keeps_slides(tmp_path, deck_with_two_slides):
    init_deck(tmp_path / "a", deck_with_two_slides)
    init_deck(tmp_path / "b", deck_with_two_slides, keep_slides=True)
    assert len(open_presentation(tmp_path / "a/a.pptx").slides) == 0
    assert len(open_presentation(tmp_path / "b/b.pptx").slides) == 2

def test_init_refuses_existing(tmp_path, potx_file):
    init_deck(tmp_path / "x", potx_file)
    with pytest.raises(PptxSkillsError, match="--reuse or --replace"):
        init_deck(tmp_path / "x", potx_file)

def test_init_never_writes_source(tmp_path, deck_with_two_slides):
    before = file_hash(deck_with_two_slides)
    init_deck(tmp_path / "y", deck_with_two_slides, keep_slides=True)
    assert file_hash(deck_with_two_slides) == before

def test_changed_externally(tmp_path, potx_file):
    init_deck(tmp_path / "z", potx_file); assert not changed_externally(tmp_path / "z")
    with open(tmp_path / "z/z.pptx", "ab") as f: f.write(b" ")
    assert changed_externally(tmp_path / "z")
```

- [ ] **Step 2: Run** — expected: FAIL.
- [ ] **Step 3: Implement.** Removing slides relies on python-pptx saving only reachable parts, so orphaned media disappear on save.
- [ ] **Step 4: Run** — expected: PASS.
- [ ] **Step 5: Commit** `git commit -m "feat(pptx): create deck folders and working copies with init"`

---

### Task 6: Change-set engine and slide operations

**Files:**
- Create: `references/pptx_skills/changeset.py`, `references/pptx_skills/schemas/changeset.schema.json`, `skills/slide-design/references/changeset-format.md` (sections: envelope, slide refs, target refs, colors, slide ops); Modify: `slides.py`, `cli.py`
- Test: `references/tests/test_changeset.py`; conftest fixture `deck_dir_from_default` (`init_deck(tmp/"decks/d", default_pptx)`, working copy `d.pptx`)

**Interfaces:**
- Consumes: `open_presentation`, `save_presentation`, `find_deck_dir`, `record_apply`, `changed_externally`, `lock_file`, `load_brief`.
- Produces:
  - Change set envelope: `{"label": "kebab-text", "ops": [ {...}, ... ]}`. Slide ref: `int` (1-based number at the time the op runs), `"id:<n>"`, or `"$last"` (slide created by the latest add/clone op in this change set). Target ref: `{"role": "title"}`, `{"name": "..."}` or `{"id": 12}`.
  - `slides.resolve_slide(prs, ref, last: Slide | None) -> Slide`; `slides.add_slide(prs, layout_name: str, after: Slide | None) -> Slide`; `slides.move_slide(prs, slide, to: int) -> None`; `slides.delete_slide(prs, slide) -> None`; `slides.layout_named(prs, name) -> SlideLayout` (error lists available names).
  - `changeset.Context` dataclass: `prs, deck_path, deck_dir, brief, examples_path, last_slide, touched: set[int], messages: list[str]`.
  - `changeset.op(name: str)` decorator registering `fn(ctx: Context, spec: dict) -> None` in `OPS`.
  - `changeset.resolve_shape(slide, target: dict) -> BaseShape` (searches group children too).
  - `changeset.apply_changeset(deck_path: Path, changeset: dict) -> dict` returning `{"seq", "touched": [slide ids], "messages": [...]}`; validates every op against the schema before changing anything; on any failure nothing is written and the error names the op index.
  - Ops in this task: `add_slide {layout, after?}`, `delete_slide {slide}`, `move_slide {slide, to}`.
  - CLI: `pptx apply <deck.pptx> <changeset.json>`.

- [ ] **Step 1: Write the failing tests**

```python
def test_add_and_move_slides(deck_dir_from_default):
    deck = deck_dir_from_default / "d.pptx"
    res = apply_changeset(deck, {"label": "two", "ops": [
        {"op": "add_slide", "layout": "Title Only"}, {"op": "add_slide", "layout": "Blank"},
        {"op": "move_slide", "slide": "$last", "to": 1}]})
    prs = open_presentation(deck)
    assert [s.slide_layout.name for s in prs.slides] == ["Blank", "Title Only"]
    assert res["seq"] == 1 and (deck_dir_from_default / "changes/001-two.json").exists()
    assert (deck_dir_from_default / "snapshots/001.pptx").exists()

def test_atomic_on_error(deck_dir_from_default):
    deck = deck_dir_from_default / "d.pptx"; before = file_hash(deck)
    with pytest.raises(PptxSkillsError, match="op 2"):
        apply_changeset(deck, {"label": "bad", "ops": [{"op": "add_slide", "layout": "Title Only"}, {"op": "delete_slide", "slide": 9}]})
    assert file_hash(deck) == before

def test_schema_rejects_unknown_op(deck_dir_from_default):
    with pytest.raises(PptxSkillsError, match="op 1.*frobnicate"):
        apply_changeset(deck_dir_from_default / "d.pptx", {"label": "x", "ops": [{"op": "frobnicate"}]})

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
```

- [ ] **Step 2: Run** — expected: FAIL.
- [ ] **Step 3: Implement** engine, slide ops, schema (one `oneOf` branch per op, `additionalProperties: false`), CLI, and the first sections of `changeset-format.md` with one worked example per op.
- [ ] **Step 4: Run** — expected: PASS.
- [ ] **Step 5: Commit** `git commit -m "feat(pptx): apply validated change sets atomically with slide operations"`

---

### Task 7: Text operations

**Files:**
- Create: `references/pptx_skills/ops_text.py`; Modify: schema, `changeset-format.md` (text section)
- Test: `references/tests/test_ops_text.py`

**Interfaces:**
- Consumes: `op`, `resolve_slide`, `resolve_shape`, `parse_color`, `apply_color`.
- Produces:
  - Op `set_text {slide, target, text? | paragraphs?, font?}`. `text` is shorthand: `\n` separates paragraphs, `**…**` marks bold runs. `paragraphs`: `[{"level"?: int, "bullet"?: bool, "runs": [{"t", "b"?, "i"?, "color"?, "size"?}]}]`. `font`: `{"size"?, "color"?, "bold"?}` applied to all runs.
  - `ops_text.write_paragraphs(text_frame, paragraphs: list[dict]) -> None` (reused by Tasks 8 and 16).

- [ ] **Step 1: Write the failing tests**

```python
def test_shorthand_bold(deck_with_title_only_slide):
    apply(deck, [{"op": "set_text", "slide": 1, "target": {"role": "title"}, "text": "**Bold** lead\nSecond"}])
    p = title_shape().text_frame.paragraphs
    assert [r.text for r in p[0].runs] == ["Bold", " lead"] and p[0].runs[0].font.bold and p[1].text == "Second"

def test_keeps_formatting_not_specified(deck_with_formatted_box):
    # box has run size 12pt, color accent1 set by hand
    apply(deck, [{"op": "set_text", "slide": 1, "target": {"name": "box"}, "text": "New"}])
    run = box().text_frame.paragraphs[0].runs[0]
    assert run.font.size.pt == 12 and read_color(run._r.rPr.find(qn("a:solidFill"))) == "accent1"

def test_paragraph_runs_with_color_and_level(deck_with_title_only_slide): ...
    # runs [{"t":"R4","color":"accent1","b":true},{"t":" Review"}], level 1 → asserted on XML

def test_target_kicker_role(kicker_deck):
    apply(deck, [{"op": "set_text", "slide": 1, "target": {"role": "kicker"}, "text": "CLASSIFICATION"}])
    assert kicker_shape().text_frame.text == "CLASSIFICATION"
```

- [ ] **Step 2: Run** — FAIL. **Step 3: Implement** — keep the first run's `rPr` as the template for new runs when the op gives no formatting. **Step 4: Run** — PASS.
- [ ] **Step 5: Commit** `git commit -m "feat(pptx): set text with runs, levels and preserved formatting"`

---

### Task 8: Shape operations

**Files:**
- Create: `references/pptx_skills/ops_shapes.py`; Modify: `geometry.py` (snapping), schema, `changeset-format.md` (shapes section)
- Test: `references/tests/test_ops_shapes.py`

**Interfaces:**
- Consumes: `op`, `resolve_*`, `parse_color`, `apply_color`, `write_paragraphs`, `load_brief`, `Box`.
- Produces:
  - `geometry.snap_box(box: Box, xs: list[float], ys: list[float], tol: float) -> Box` — snaps each edge to the nearest line within `tol`; lines = content-area edges and guides from the brief.
  - Op `add_shape {slide, name, shape, box, fill?, line?, pattern?, text?|paragraphs?, font?, align?, anchor?, insets?, radius?, allow_rgb?}`; `shape` ∈ `rect roundRect homePlate chevron rightArrow leftArrow upArrow downArrow ellipse textbox`; `box` `[x,y,w,h]` cm; `line` `"none"` or `{color, width_pt, dash?}`; `pattern` `{preset, fg, bg}` (OOXML preset names, e.g. `wdUpDiag`); `align` `left|center|right`; `anchor` `top|middle|bottom`; `insets` `[l,t,r,b]` cm; `radius` 0–0.5 (roundRect adj).
  - Op `add_line {slide, name, from [x,y], to [x,y], color, width_pt, dash?, head?, tail?}` (`head`/`tail` ∈ `none triangle arrow`).
  - Ops `update_shape {slide, target, …any add_shape field except shape}`, `delete_shape {slide, target}`, `z_order {slide, target, to: front|back|forward|backward}`, `group {slide, targets: [...], name}`, `ungroup {slide, target}`.
  - Every added shape gets `name` exactly as given; duplicate names on a slide → error.

- [ ] **Step 1: Write the failing tests**

```python
def test_add_shape_theme_fill_text_and_snap(deck_with_guides):   # content area left edge 1.5
    apply(deck, [{"op": "add_shape", "slide": 1, "name": "cell-r1", "shape": "rect", "box": [1.53, 5, 8, 2],
                  "fill": "accent1+60", "line": "none", "text": "**Kein Ziel**\nLohnt sich nicht", "font": {"size": 12}}])
    s = shape_named("cell-r1")
    assert inv_shape("cell-r1")["box"][0] == 1.5 and inv_shape("cell-r1")["fill"] == "accent1+60"
    assert s.text_frame.paragraphs[0].runs[0].font.bold

def test_pattern_fill(deck):
    apply(deck, [{"op": "add_shape", ..., "pattern": {"preset": "wdUpDiag", "fg": "accent6", "bg": "lt1"}}])
    assert inv_shape("hatch")["fill"] == "pattern:wdUpDiag accent6 lt1"

def test_rgb_requires_flag(deck):
    with pytest.raises(PptxSkillsError, match="allow_rgb"):
        apply(deck, [{"op": "add_shape", ..., "fill": "#FF0000"}])

def test_line_with_arrowhead(deck):
    apply(deck, [{"op": "add_line", "slide": 1, "name": "path", "from": [10, 8], "to": [14, 8], "color": "dk1", "width_pt": 3, "tail": "triangle"}])
    assert shape_named("path")._element.find(".//" + qn("a:tailEnd")).get("type") == "triangle"

def test_update_delete_zorder_group_ungroup(deck_with_two_boxes):
    apply(deck, [{"op": "update_shape", "slide": 1, "target": {"name": "a"}, "box": [2, 2, 3, 3]},
                 {"op": "z_order", "slide": 1, "target": {"name": "a"}, "to": "front"},
                 {"op": "group", "slide": 1, "targets": [{"name": "a"}, {"name": "b"}], "name": "ab"},
                 {"op": "ungroup", "slide": 1, "target": {"name": "ab"}},
                 {"op": "delete_shape", "slide": 1, "target": {"name": "b"}}])
    assert inv_shape("a")["box"] == [2.0, 2.0, 3.0, 3.0] and not exists("b") and not exists("ab")
```

- [ ] **Step 2: Run** — FAIL. **Step 3: Implement.** Ungroup: move children to the group's parent, converting each child `xfrm` to absolute coordinates with the group's `off/ext/chOff/chExt`. Arrowheads via `a:headEnd`/`a:tailEnd` on the line's `a:ln`. **Step 4: Run** — PASS.
- [ ] **Step 5: Commit** `git commit -m "feat(pptx): add, change, group and remove shapes with theme styling"`

---

### Task 9: Grid command

**Files:**
- Create: `references/pptx_skills/grid.py`; Modify: `cli.py`, `changeset-format.md` (grid tip)
- Test: `references/tests/test_grid.py`

**Interfaces:**
- Consumes: `Box`, `load_brief`.
- Produces: `grid_boxes(area: Box, cols: int, rows: int = 1, gutter: float = 0.3, row_gutter: float | None = None) -> list[list[Box]]` (row-major; the last column/row absorbs rounding so outer edges equal the area's edges exactly); CLI `pptx grid <deck-dir> --cols C [--rows R] [--gutter G] [--row-gutter G] [--area content|full|x,y,w,h] [--json]`.

- [ ] **Step 1: Write the failing test**

```python
def test_grid_four_columns():
    rows = grid_boxes(Box(1.5, 4.82, 30.87, 11.88), cols=4, gutter=0.3)
    assert len(rows) == 1 and len(rows[0]) == 4
    assert rows[0][0].x == 1.5 and rows[0][3].right == pytest.approx(32.37, abs=0.001)
    assert rows[0][1].x - rows[0][0].right == pytest.approx(0.3, abs=0.01)
```

- [ ] **Step 2–4:** run (FAIL), implement, run (PASS). **Step 5: Commit** `git commit -m "feat(pptx): compute aligned grid boxes inside the content area"`

---

### Task 10: Fonts and text measurement

**Files:**
- Create: `references/pptx_skills/fonts.py`
- Test: `references/tests/test_fonts.py`

**Interfaces:**
- Consumes: `load_brief` data (theme fonts), python-pptx slide objects.
- Produces:
  - `RunStyle(family: str, size_pt: float, bold: bool, italic: bool)`.
  - `effective_run_style(run_el, paragraph_el, shape, slide) -> RunStyle`.
  - `find_font(family: str, bold: bool = False, italic: bool = False) -> Path | None` — searches `/System/Library/Fonts`, `/System/Library/Fonts/Supplemental`, `/Library/Fonts`, `~/Library/Fonts`, `/usr/share/fonts`, `~/.fonts` (results cached per process), matching `ImageFont.truetype(path).getname()`.
  - `measure_paragraphs(shape, slide) -> Measurement(lines_per_paragraph: list[int], height_cm: float, widest_word_cm: float, fonts_missing: set[str])`.

Effective style resolution order (first explicit value wins): run `rPr` → paragraph `pPr/defRPr` → shape `txBody/lstStyle` level → (placeholders) layout placeholder `lstStyle` → master placeholder `lstStyle` → master `txStyles` (`titleStyle` for title/ctrTitle, `bodyStyle` for other placeholders, `otherStyle` for non-placeholders) → presentation `defaultTextStyle` → 18 pt. Theme font tokens `+mj-lt`/`+mn-lt` resolve to the theme's major/minor font; title placeholders default to major.

Measurement: greedy word wrap per paragraph at the box width minus insets (defaults 0.25 cm left/right, 0.13 cm top/bottom), widths via Pillow `FreeTypeFont.getlength` at 10× size for precision; `a:br` forces a line break; line height = largest run size on the line × 1.2 × `lnSpc` percentage; add `spcBef`/`spcAft`. When a font is missing, measure with `ImageFont.load_default(size)` and add the family to `fonts_missing`.

- [ ] **Step 1: Write the failing tests**

```python
def test_effective_style_inherits_from_master(default_pptx):
    prs = Presentation(default_pptx); s = prs.slides.add_slide(prs.slide_layouts[1]); s.placeholders[1].text = "x"
    p = s.placeholders[1].text_frame.paragraphs[0]
    st = effective_run_style(p.runs[0]._r, p._p, s.placeholders[1], s)
    assert (st.family, st.size_pt) == ("Calibri", 32)

def test_wrap_counts_lines_with_default_font(slide_with_box):
    box = textbox(width_cm=5, text="word " * 40, size=12)
    m = measure_paragraphs(box, slide)
    assert m.lines_per_paragraph[0] > 3 and m.height_cm > 2

def test_hard_break_and_missing_font(slide_with_box):
    box = textbox(width_cm=20, text="a\vb", size=12, family="NoSuchFont")
    m = measure_paragraphs(box, slide)
    assert m.lines_per_paragraph == [2] and "NoSuchFont" in m.fonts_missing
```

- [ ] **Step 2–4:** run (FAIL), implement, run (PASS). **Step 5: Commit** `git commit -m "feat(pptx): resolve effective text styles and measure wrapped text"`

---

### Task 11: Checks — structure and text rules, report, `check` command

**Files:**
- Create: `references/pptx_skills/checks/{__init__,structure,text}.py`; Modify: `cli.py`
- Test: `references/tests/test_checks_text.py`

**Interfaces:**
- Consumes: `inventory`, `measure_paragraphs`, `effective_run_style`, `load_brief`, `load_deck`, `placeholder_role`.
- Produces:
  - `Finding(rule, category, severity, slide, slide_id, shape_id, shape, message, fix)` dataclass; categories `structure text geometry style`; severities `error warning info`.
  - `run_checks(deck_path: Path, slides: list[int] | None = None, all_slides: bool = False) -> dict` → `{"deck", "slides": [numbers], "findings": [...], "summary": {"error": n, "warning": n, "info": n}}`. Default slide set: `last_touched` from deck.json, else all slides.
  - Rule registry: `checks.rule(rule_id, category, severity)` decorator; rules take `(ctx, slide) -> list[Finding]`.
  - Rules here: `reopen`, `relationships`, `duplicate-ids` (structure, error); `overflow` (error), `leftover-text` (error: text equal to the layout placeholder's prompt, or matching `\bLorem\b`, `\bTODO\b`, `\bXXX\b`), `min-size` (error: below `min_size_pt`; `min_size_small_pt` applies to footer/number/date placeholders and shapes whose top is below the content area's bottom), `too-many-lines` (warning: paragraph wraps to > `max_lines`, title > `max_title_lines`), `too-many-bullets` (warning: > `max_bullets` bullet paragraphs in a shape), `keyword-chain` (warning), `shrink-to-fit` (warning: `normAutofit` with `fontScale`), `empty-placeholder` (warning; not footer/number/date), `topic-title` (info: title < `title_min_words` words).
  - CLI: `pptx check <deck> [--slides 1,2] [--all] [--json] [--strict]`; exit 1 when any error (with `--strict`, also on warnings); writes `check/slide-NN.json` per checked slide when in a deck folder.

- [ ] **Step 1: Write the failing tests**

```python
def test_overflow_is_error(deck): add box 5×1 cm with 60 words at 12 pt
    assert finding(run_checks(deck), "overflow")["severity"] == "error"

def test_keyword_chain_only_for_three_items(deck):
    texts = {"chain": "Speed · Quality · Scale", "footer": "|   Company AG", "eg": "z. B. Pipelines, Gates, Audit-Trail", "two": "A | B"}
    flagged = {f["shape"] for f in run_checks(deck)["findings"] if f["rule"] == "keyword-chain"}
    assert flagged == {"chain"}

def test_leftover_prompt_text(deck_from_kicker_master):   # kicker placeholder left with its prompt
    assert finding(run_checks(deck), "leftover-text")["shape_id"] == kicker_id

def test_min_size_and_small_zone(deck): body run 10 pt → error; same run in a box below the content area → none
def test_lines_per_paragraph_not_per_shape(deck): one shape with 4 one-line paragraphs → no too-many-lines; one 5-line paragraph → warning
def test_exit_codes(deck_with_overflow, clean_deck):
    assert main(["check", str(deck_with_overflow)]) == 1 and main(["check", str(clean_deck)]) == 0
```

- [ ] **Step 2–4:** run (FAIL), implement, run (PASS). **Step 5: Commit** `git commit -m "feat(pptx): check structure and text rules with a JSON report"`

---

### Task 12: Render, contact sheets and `doctor`

**Files:**
- Create: `references/pptx_skills/render.py`; Modify: `brief.py` (`style` command writes contact sheets), `cli.py`
- Test: `references/tests/test_render.py`

**Interfaces:**
- Consumes: `open_presentation`, `save_presentation`, `slides.add_slide`, `find_deck_dir`, `find_font`.
- Produces:
  - `find_soffice() -> Path | None` (PATH, then `/Applications/LibreOffice.app/Contents/MacOS/soffice`); `find_pdftoppm() -> Path | None`.
  - `render(deck: Path, out_dir: Path | None = None, slides: list[int] | None = None, dpi: int = 120) -> list[Path]` — `out_dir` defaults to `<deck-dir>/render/` or `<deck-stem>-render/`; files `slide-NN.png` (NN = slide number, two digits). Runs `soffice --headless -env:UserInstallation=file://<tmp-profile> --convert-to pdf --outdir <tmp> <deck>` (timeout 180 s) then `pdftoppm -png -r <dpi> -f N -l N`. Missing tools → `PptxSkillsError` naming `brew install --cask libreoffice` / `brew install poppler` (and `apt install libreoffice poppler-utils`).
  - `contact_sheet(pngs: list[Path], labels: list[str], out: Path, columns: int = 4, thumb_width: int = 480) -> Path`.
  - `render_layouts(deck: Path, out: Path) -> Path` — temporary deck with one slide per layout, every placeholder filled with its role name, labeled with the layout name → `style/layouts.png`; `style` also writes `style/examples.png` for the examples deck. `style` stores `sources_hash` (sha256 over the source files) in `brief.json` and skips the contact sheets when it is unchanged, unless `--refresh`.
  - `doctor(deck_dir: Path | None) -> dict` → `{"soffice": path|None, "pdftoppm": path|None, "fonts": {family: path|None}}`; CLI `pptx doctor [deck-dir]`, exit 1 if anything is missing.
  - CLI: `pptx render <deck> [--slides 1,3] [--dpi N] [--sheet]`.

- [ ] **Step 1: Write the failing tests**

```python
def test_render_invokes_tools_with_isolated_profile(monkeypatch, deck):
    calls = record_subprocess(monkeypatch, fake_pdf_and_pngs)
    out = render(deck, slides=[2])
    assert any("-env:UserInstallation=file://" in a for a in calls[0])
    i = calls[1].index("-f"); assert calls[1][i:i + 4] == ["-f", "2", "-l", "2"]
    assert [p.name for p in out] == ["slide-02.png"]

def test_missing_soffice_names_install_command(monkeypatch, deck):
    monkeypatch.setattr(render_mod, "find_soffice", lambda: None)
    with pytest.raises(PptxSkillsError, match="brew install --cask libreoffice"): render(deck)

def test_contact_sheet_size(tmp_path): 5 PNGs 1600×900, columns 4 → sheet width 4*480 + gaps, 2 rows

@pytest.mark.skipif(find_soffice() is None or find_pdftoppm() is None, reason="LibreOffice/poppler not installed")
def test_render_real(deck_with_two_slides, tmp_path):
    assert [p.name for p in render(deck_with_two_slides, tmp_path)] == ["slide-01.png", "slide-02.png"]
```

- [ ] **Step 2–4:** run (FAIL), implement, run (PASS). Then install the tools (`brew install --cask libreoffice && brew install poppler`) and run the real test — expected: PASS.
- [ ] **Step 5: Commit** `git commit -m "feat(pptx): render slides and contact sheets with LibreOffice and poppler"`

---

### Task 13: `slide-design` skill (Milestone M1)

**Files:**
- Create: `skills/slide-design/SKILL.md`, `skills/slide-design/references/visual-forms.md`, `skills/slide-design/references/review-checklist.md`; Modify: `changeset-format.md` (complete for Tasks 6–9); delete `skills/.gitkeep`
- Test: `references/tests/test_skills.py`

Write with `superpowers:writing-skills`. Content (from spec §4.3, §4.4, §8, §9):
- Frontmatter `description`: what it does and "Use when …" with English and German triggers ("design a slide", "Folie gestalten", "make this slide …", a screenshot plus change request, "Folie zu …"); state it should be preferred over generic pptx skills when a master or `decks/` folder is involved.
- Boundary table `slide-design` vs `deck-storyline`.
- Setup: how to locate `<skill-dir>` and run `uv run --project <skill-dir>/../../references pptx …`; first run `pptx doctor`.
- Workflow steps 1–7 of spec §4.3, the screenshot procedure of §4.4, getting slides back (§4.5), the lock-file note.
- Design rules 1–12 and the anti-pattern list of spec §8 (the method; mark Volker's two as non-negotiable).
- The render-check-review loop with the 3-round limit and what to report.
- Pointers: `changeset-format.md`, `visual-forms.md` (matrix, process/chevrons, comparison, timeline, tiles, chart, table — each with when to use and a `pptx grid` recipe), `review-checklist.md` (spec §9 visual list).

- [ ] **Step 1: Write the failing tests**

```python
SKILLS = sorted((ROOT / "skills").glob("*/SKILL.md"))

def test_frontmatter_and_length():
    for md in SKILLS:
        fm = frontmatter(md)
        assert set(fm) == {"name", "description"} and fm["name"] == md.parent.name
        assert len(md.read_text().splitlines()) < 300

def test_commands_in_skills_exist():
    for md in [*SKILLS, *(ROOT / "skills").glob("*/references/*.md")]:
        for cmd in set(re.findall(r"\bpptx (\w+)", md.read_text())):
            assert main([cmd, "--help"]) == 0, cmd

def test_no_agent_specific_tool_names():
    for md in [*SKILLS, *(ROOT / "skills").glob("*/references/*.md")]:
        assert not re.search(r"\b(Read|Write|Edit|Bash|Grep|Glob) tool\b", md.read_text())
```

- [ ] **Step 2: Run** — FAIL (no SKILL.md). **Step 3: Write** the skill files. **Step 4: Run** — PASS.
- [ ] **Step 5: Smoke test M1** (manual, local): `claude --plugin-dir ./powerpoint-skills`, ask for one slide on the corporate master (`init` from the local `.potx`), check that it renders, checks and shows. Note problems in the task log.
- [ ] **Step 6: Commit** `git commit -m "feat(slide-design): add the slide-by-slide design skill"`

---

### Task 14: Checks — geometry and style rules

**Files:**
- Create: `references/pptx_skills/checks/{geometry,style}.py`
- Test: `references/tests/test_checks_layout.py`

**Interfaces:**
- Consumes: rule registry, `absolute_box`, brief (`content_area`, `rules`, theme fonts), `find_font`.
- Produces rules: `off-slide` (error, tolerance 0.01 cm), `outside-content` (warning; excludes title/kicker/footer/number/date placeholders), `overlap` (warning: two text-bearing shapes intersect by > 0.01 cm² and neither contains the other; lines/connectors excluded), `near-aligned` (warning: left/right/top/bottom edges of two shapes differ by 0 < d ≤ `near_align_cm`), `uneven-gaps` (warning: ≥ 3 shapes in a row — vertical overlap ≥ 50 % — whose consecutive gaps (all < 3 cm) differ by > 0.1 cm), `rgb-color` (warning: `a:srgbClr` in the slide's own shape XML), `foreign-font` (warning: explicit `latin` typeface not a theme font or theme token), `missing-font` (warning, once: theme font not found by `find_font`), `alt-text` (info: picture without `descr`).

- [ ] **Step 1: Write the failing tests**

```python
def test_off_slide_child_in_group(deck): group whose child extends past the right slide edge → "off-slide" error on the child
def test_overlap_vs_containment(deck): box inside box → none; two crossing text boxes → "overlap"
def test_near_aligned(deck): left edges 1.50 and 1.70 → warning; 1.50 and 1.50 → none; 1.50 and 2.00 → none
def test_uneven_gaps(deck): three boxes with gaps 0.3 and 0.6 → warning
def test_rgb_rule_ignores_scheme_aliases(deck): fills tx1, bg1+10, accent1 → none; one srgbClr fill → one warning
```

- [ ] **Step 2–4:** run (FAIL), implement, run (PASS). **Step 5: Commit** `git commit -m "feat(pptx): check geometry, alignment, colors and fonts"`

---

### Task 15: Clone and replace slides across decks

**Files:**
- Modify: `references/pptx_skills/slides.py`, schema, `changeset-format.md`
- Test: `references/tests/test_clone.py`

**Interfaces:**
- Consumes: `layout_named`, `resolve_slide`, `op`, deck.json sources (`examples`).
- Produces:
  - `clone_slide(src_prs, src_slide, dst_prs, after: Slide | None) -> Slide` and `replace_slide(dst_prs, target: Slide, src_prs, src_slide) -> Slide` (keeps the target's position and `sldId` id).
  - Ops `clone_slide {from: {deck: "examples"|"self"|"<path>", slide}, after?}` (`"self"` = duplicate) and `replace_slide {slide, from: {deck, slide}}`.

Clone algorithm: map the source slide's layout to the destination layout of the same name (else `PptxSkillsError("layout '<name>' not in working copy; available: …")`); add a slide on it and drop its auto-created placeholders; deep-copy the source `spTree` children and `p:bg`; for every `r:embed`/`r:link`/`r:id` in the copy, re-create the relationship: image → `get_or_add_image_part`; external hyperlink → `relate_to(url, RT.HYPERLINK, is_external=True)`; chart → load a new `ChartPart` from the source blob with `c:externalData` removed, then `chart_workbook.update_from_xlsx_blob(src_workbook_blob)`; `tags` → drop `p:custDataLst` and add a message; any other type (`oleObject`, media, diagrams) → `PptxSkillsError("slide N uses <type>, which cannot be cloned")`. Copy notes text.

- [ ] **Step 1: Write the failing tests**

```python
def test_clone_with_picture_and_chart(examples_deck, deck_dir):   # generated: slide with PNG + column chart
    apply(deck, [{"op": "clone_slide", "from": {"deck": "examples", "slide": 1}}])
    prs = open_presentation(deck); s = prs.slides[0]
    assert {sh.shape_type for sh in s.shapes} >= {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART}
    assert s.shapes[chart_idx].chart.plots[0].categories[0] == "A"
    assert run_checks(deck, all_slides=True)["summary"]["error"] == 0

def test_replace_keeps_position_and_id(deck_with_three_slides, other_deck):
    old_id = open_presentation(deck).slides[1].slide_id
    apply(deck, [{"op": "replace_slide", "slide": 2, "from": {"deck": str(other_deck), "slide": 1}}])
    s = open_presentation(deck).slides[1]
    assert s.slide_id == old_id and s.shapes.title.text == "From other deck"

def test_missing_layout_and_ole(deck, deck_with_unknown_layout, deck_with_ole): errors match "not in working copy" / "oleObject"
def test_duplicate_self(deck): clone from "self" slide 1 → 2 slides, identical text
```

- [ ] **Step 2–4:** run (FAIL), implement, run (PASS). **Step 5: Commit** `git commit -m "feat(pptx): clone and replace slides across decks"`

---

### Task 16: Tables

**Files:**
- Create: `references/pptx_skills/ops_tables.py`; Modify: schema, `changeset-format.md`, `inventory.py` (table cell texts as `"table": [[...]]`)
- Test: `references/tests/test_ops_tables.py`

**Interfaces:**
- Produces: op `add_table {slide, name, box, rows: [[cell text…]], header?: bool, style_id?, col_widths?: [cm], font?: {size, color}, fills?: {"r,c": color}}`; op `set_table {slide, target, rows}` — sets cell texts, adds rows by deep-copying the last `a:tr`, removes surplus rows; a column-count mismatch is an error. Cell text uses the `text` shorthand of Task 7.

- [ ] **Step 1: Write the failing tests**

```python
def test_add_table(deck): 3×2 table → inventory "table" == [["A","B"],["1","2"],["3","4"]], box as given
def test_set_table_grows_and_shrinks(deck_with_cloned_table): 2 → 4 rows keeps the last row's formatting; 4 → 1 row
def test_column_mismatch(deck_with_table): rows with 3 columns into a 2-column table → error mentions "columns"
```

- [ ] **Step 2–4:** run (FAIL), implement, run (PASS). **Step 5: Commit** `git commit -m "feat(pptx): add and fill native tables"`

---

### Task 17: Charts, pictures and notes

**Files:**
- Create: `references/pptx_skills/ops_media.py`; Modify: schema, `changeset-format.md`
- Test: `references/tests/test_ops_media.py`

**Interfaces:**
- Produces:
  - Op `add_chart {slide, name, box, type, categories, series: [{name, values}], legend?: bottom|right|none, labels?: bool}`; `type` ∈ `column_clustered column_stacked bar_clustered bar_stacked line doughnut`; series colored `accent1…accent6` in order via theme colors; fonts from the theme at `min_size_pt` or larger.
  - Op `add_picture {slide, name, path, box? | target?, alt}` — PNG/JPG only; into a picture placeholder via `insert_picture`, else into `box` cropped to fill (keep aspect ratio via `crop_*`); `alt` required.
  - Op `set_notes {slide, text}`.

- [ ] **Step 1: Write the failing tests**

```python
def test_column_chart_editable(deck): chart type COLUMN_CLUSTERED, categories ["A","B"], series values (1.0, 2.0), first series fill theme accent1
def test_picture_cropped_to_fill(deck, png_200x100): box 4×4 cm → crop_left + crop_right ≈ 0.5, crop_top == 0; descr == alt
def test_picture_rejects_svg(deck, svg_file): error mentions "PNG or JPG"
def test_notes(deck): notes_slide text == "Say this"
```

- [ ] **Step 2–4:** run (FAIL), implement, run (PASS). **Step 5: Commit** `git commit -m "feat(pptx): add native charts, pictures and speaker notes"`

---

### Task 18: Neutral template and style file

**Files:**
- Create: `references/pptx_skills/neutral.py`, `references/pptx_skills/schemas/style-file.schema.json`; Modify: `deckdir.py` (style-file decks; `--style-file` rules merged into the brief for any source)
- Test: `references/tests/test_neutral.py`

**Interfaces:**
- Produces:
  - Style file: `{"colors": {role: "#RRGGBB"}, "fonts": {"major", "minor"}, "rules": {…subset of DEFAULT_RULES}}`; all keys optional.
  - `NEUTRAL_COLORS = {"dk1": "#000000", "lt1": "#FFFFFF", "dk2": "#333333", "lt2": "#EDEDED", "accent1": "#1F4E79", "accent2": "#2E75B6", "accent3": "#7F7F7F", "accent4": "#BFBFBF", "accent5": "#D9D9D9", "accent6": "#C55A11", "hlink": "#1F4E79", "folHlink": "#7F7F7F"}`; default fonts Arial/Arial.
  - `neutral_presentation(style: dict) -> Presentation` — python-pptx default template converted to 16:9 (12192000 × 6858000 EMU; x offsets and widths of all master/layout shapes scaled by 4/3), theme colors written as `srgbClr` (replacing `sysClr`), font scheme set.

- [ ] **Step 1: Write the failing tests**

```python
def test_neutral_is_16_9_with_style(tmp_path):
    prs = neutral_presentation({"colors": {"accent1": "#123456"}, "fonts": {"major": "Georgia", "minor": "Verdana"}})
    b = build_brief(prs)
    assert b["slide_size_cm"] == [33.87, 19.05]
    assert b["theme"]["colors"]["accent1"] == "#123456" and b["theme"]["colors"]["accent2"] == "#2E75B6"
    assert b["theme"]["fonts"] == {"major": "Georgia", "minor": "Verdana"}
    assert b["content_area"] == pytest.approx([1.69, 4.45, 30.48, 12.57], abs=0.02)

def test_init_with_style_file_only(tmp_path, style_json): init → working copy with 0 slides and the style's accent1
def test_style_file_rules_apply_to_any_source(tmp_path, potx_file, style_json_with_rules): brief rules min_size_pt == 10
def test_invalid_style_file(tmp_path): {"colors": {"accent9": "#000000"}} → error names "accent9"
```

- [ ] **Step 2–4:** run (FAIL), implement, run (PASS). **Step 5: Commit** `git commit -m "feat(pptx): support style files with a generated neutral template"`

---

### Task 19: `deck-storyline` skill

**Files:**
- Create: `skills/deck-storyline/SKILL.md`, `skills/deck-storyline/references/storyline-format.md`; Modify: `inventory.py` (`--outline`: one line per slide with number, id, kicker, title), `skills/slide-design/SKILL.md` (status updates and slide ids in `storyline.md`)
- Test: `references/tests/test_skills.py` (already covers both skills), `references/tests/test_inventory.py::test_outline`

Write with `superpowers:writing-skills`. Content (spec §4.2): inputs (text, topic, files or a research run folder), the few questions to ask (audience, live vs pre-read, length, language), the storyline header and per-slide fields, statuses `open / in design / approved / existing`, slide ids, continuing an existing deck via `pptx inventory --outline`, content rules 1–6 of spec §8 applied to titles, hand-over to `slide-design`. `storyline-format.md` holds the exact markdown format with one full example (generic topic, no company).

- [ ] **Step 1: Write the failing test** `test_outline`: deck with two titled slides → `main(["inventory", deck, "--outline"])` prints `1 (id 256) | <kicker> | First title` and `2 (id 257) | … | Second title`.
- [ ] **Step 2–4:** run (FAIL), implement and write the skill, run `test_inventory.py` and `test_skills.py` (PASS).
- [ ] **Step 5: Commit** `git commit -m "feat(deck-storyline): add the storyline skill"`

---

### Task 20: Packaging — README, manifests, Codex, repo hygiene

**Files:**
- Modify: `README.md`, `.claude-plugin/plugin.json` (description)
- Create: `scripts/sync-mirrors.sh`, `.codex-plugin/plugin.json` (generated), `references/tests/test_repo.py`

**Interfaces:**
- Produces:
  - `plugin.json` description (and later the marketplace entry): "Design PowerPoint slides one at a time in the style of any master or deck: a storyline first, then each slide built natively in a working copy, rendered to PNG and checked before you see it."
  - `scripts/sync-mirrors.sh [--check]` generating only `.codex-plugin/plugin.json` from `plugin.json` (`name, version, description, author, homepage, repository, license, "skills": "./skills/"`), same style as `research-toolkit/scripts/sync-mirrors.sh`.
  - README sections: Skills (both, with the boundary table), Installation (marketplace lines; Codex), Setup (`brew install --cask libreoffice`, `brew install poppler`, `cd references && uv sync`, `uv run pptx doctor`; Linux equivalents), Customization (style sources, style file, rules overrides), Output Convention (`decks/<slug>/…` tree from spec §5.3), File Structure.

- [ ] **Step 1: Write the failing tests**

```python
def test_no_office_files_tracked():
    files = subprocess.run(["git", "ls-files"], cwd=ROOT, capture_output=True, text=True).stdout.split()
    assert not [f for f in files if f.lower().endswith((".pptx", ".potx", ".pptm", ".ppt", ".odp", ".png", ".jpg"))]

@pytest.mark.skipif(not os.environ.get("PPTX_SKILLS_DENY_WORDS"), reason="local only")
def test_no_company_names():
    words = [w.strip().lower() for w in os.environ["PPTX_SKILLS_DENY_WORDS"].split(",")]
    for f in tracked_text_files():
        assert not any(w in f.read_text(errors="ignore").lower() for w in words), f

def test_codex_manifest_in_sync():
    assert subprocess.run(["bash", "scripts/sync-mirrors.sh", "--check"], cwd=ROOT).returncode == 0
```

- [ ] **Step 2: Run** — FAIL. **Step 3: Implement** script, run it, write README and description. **Step 4: Run** all tests plus `claude plugin validate .` and `/plugin-audit powerpoint-skills` (workspace skill) — expected: green, no blockers.
- [ ] **Step 5: Commit** `git commit -m "chore: add README, Codex manifest and repo hygiene checks"`

---

### Task 21: Local acceptance tests and acceptance run

**Files:**
- Create: `references/tests/test_local_acceptance.py` (all tests skip unless the env vars are set)

**Interfaces:**
- Consumes: everything above; env vars `PPTX_SKILLS_LOCAL_MASTER`, `PPTX_SKILLS_LOCAL_EXAMPLES`, `PPTX_SKILLS_LOCAL_REFERENCE`.

- [ ] **Step 1: Write the tests** (generic — no company names or master-specific numbers):

```python
@needs("PPTX_SKILLS_LOCAL_REFERENCE")
def test_reference_slide_passes(ref):
    rep = run_checks(ref, all_slides=True)
    assert rep["summary"]["error"] == 0
    assert not [f for f in rep["findings"] if f["category"] == "text" and f["severity"] == "warning"]

@needs("PPTX_SKILLS_LOCAL_MASTER")
def test_master_brief(tmp_path, master):
    d = init_deck(tmp_path / "m", master); b = load_brief(tmp_path / "m")
    assert b["content_area_source"] in ("guides", "layouts") and any(p["role"] == "kicker" for l in b["layouts"] for p in l["placeholders"])

@needs("PPTX_SKILLS_LOCAL_MASTER", "PPTX_SKILLS_LOCAL_EXAMPLES")
def test_clone_every_example(tmp_path, master, examples):
    # clone each example slide; slides with unsupported parts must fail with a clear "cannot be cloned" error, all others pass structure checks
```

- [ ] **Step 2: Run** with the env vars pointing at the local files — expected: PASS. If the reference slide fails a text rule, fix the rule (not the threshold values in Global Constraints unless Volker agrees) and re-run.
- [ ] **Step 3: Acceptance run with Volker** (spec §14): storyline + five slides on the corporate master; reference slide rebuilt from its screenshot; screenshot feedback on a hand-edited slide keeps the edits; a second style source and a style file; one slide workflow in Codex CLI; every working copy opens in PowerPoint without a repair prompt. Record the results in the PR-less merge notes (commit body).
- [ ] **Step 4: Commit** `git commit -m "test: add local acceptance tests against private masters"`
