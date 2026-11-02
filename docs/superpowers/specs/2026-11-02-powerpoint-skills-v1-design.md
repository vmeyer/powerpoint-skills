# powerpoint-skills v1 — Design

Date: 2026-11-02 · Status: approved · Supersedes the parts of `REQUIREMENTS.md` and
`HANDOFF.md` that conflict with it (see "Changes against REQUIREMENTS.md").

## 1. Intent

Volker builds presentations the way he works with colleagues: first agree on a **storyline**,
then design the deck **one slide at a time**, look at each slide, say what is wrong, refine it,
and only then move on. In his experience this is faster and gives much better slides than
generating a whole deck in one go.

v1 brings that way of working to coding agents. The agent turns text, a topic or a prepared
research run into a storyline, then designs single slides in the style of an existing master,
renders and checks each slide itself, shows it, and revises it from feedback — including a
screenshot of the slide with a note on what to change.

The plugin is deliberately small. Volker will use it on real decks and grow it from what he
misses, instead of trying to get the complete skill right up front.

The model for the capabilities is Anthropic's built-in `pptx` skill, which Volker is happy
with. This plugin is a clean-room build (see Constraints) that adds what he needs on top:
style from his own master, the slide-by-slide workflow, and use from more than one agent.

### The quality bar

The reference is a slide Volker built himself (kept locally as `.pptx` and screenshot, never in
the repo). It uses only theme colors (no RGB values), a pattern fill for the hatching, 11–12 pt
body text (below the master's 14 pt body style) and about 250 words. It is a content layout with:

- a short uppercase **kicker** in the accent color naming the slide's role ("Classification"),
- a two-line **action title** that states the takeaway as full sentences,
- a **matrix built entirely from native shapes** on the master's grid: two labeled axes drawn as
  arrow shapes, row and column headers each with a bold label, a short explanation and
  examples, cells colored by meaning (tints of the accent color for value, a light highlight
  color for risk, a hatched area for "not sensible"), one highlight badge ("Today"), and arrows
  showing the path to the target state,
- a short lead-in block on the left with **bold lead-ins** followed by one plain sentence.

Nothing on it is a picture, every text fits, every edge sits on a shared grid, and every color
comes from the theme. That is the standard v1 slides are judged against.

## 2. Goals, non-goals, anti-goals

### Goals (v1)

1. **Storyline first.** From text, a topic or a research run (for example a
   `research-toolkit` result folder), the agent proposes a storyline: one entry per slide with
   kicker, action title and what the slide shows. Volker adjusts it in conversation. The
   storyline is a file, not just chat.
2. **One slide at a time.** The agent creates or revises exactly one slide per step, renders it,
   checks it, shows it and waits for feedback. Approval moves the storyline on to the next slide.
3. **Feedback by text or screenshot.** Volker can say what to change, or send a screenshot of an
   existing slide (marked up or not) with a note. The agent finds the slide and changes the real
   slide in the `.pptx`, never a picture of it.
4. **The `.pptx` is the source of truth.** The agent always reads the current state of the file
   before changing anything and changes only what it was asked to. Edits made by hand in
   PowerPoint survive.
5. **The agent works on its own copy.** The first step is always a local working copy; the user's
   original file is never written.
6. **Any style source.** Style comes from what the user provides, alone or combined:
   - a master or template (`.potx`, or a `.pptx` used as one),
   - an existing deck (`.pptx`), whose real slides show how the style is applied
     (cards, tables, spacing, color use) and can be cloned as building blocks,
   - a small style file (colors, fonts, rules) when there is no PowerPoint source; the agent then
     works on a neutral template with those colors and fonts.
7. **Self-check before showing.** Every changed slide is rendered to PNG, checked by
   deterministic rules (overflow, off-grid, off-slide, theme colors, leftover template text, …)
   and reviewed visually by the agent, which fixes what it finds before showing the slide —
   the same render-and-fix loop as `diagram-skills`.
8. **Usable from more than one agent.** All logic sits in shell-invoked scripts; the skills
   contain only shell commands and plain instructions. Claude Code is the primary target; Codex
   CLI is supported through a generated manifest.

### Non-goals (v1)

- Generating a complete deck in one go (the storyline exists; slides are still designed one at
  a time).
- Writing changes back into the user's original file. Volker copies finished slides into his
  deck by hand (my call; see Decisions).
- Re-branding foreign decks onto another master, merging or splitting decks, comments.
- Animations, transitions, audio, video, macros (`.pptm`), `.ppt`/`.odp` input.
- Cloud APIs (Google Slides, PowerPoint Online).
- Running checks in CI against user decks, full ECMA-376 XSD validation.
- Rendering through PowerPoint itself (LibreOffice only; see Later).
- SVG and Excalidraw import, icon libraries.
- Windows support beyond what works by accident (macOS primary, Linux should work).

### Anti-goals (would be a failure even if it technically works)

- Slides that read as generic AI output: keyword chains joined by middle dots
  ("Speed · Quality · Scale"), walls of text, topic titles instead of statements (see §8).
- Writing to the user's original file, or losing hand edits made in PowerPoint.
- **Any corporate master, example deck or screenshot ending up in this public repo.**
- An own theming system that overrides or rebuilds the master, or hard-coded colors.
- Text, diagrams or tables pasted in as pictures instead of native, editable shapes.

## 3. Constraints

- **Public repo.** `vmeyer/powerpoint-skills` is public. Corporate masters, example decks and
  screenshots stay local. The repo-wide `.gitignore` excludes `*.pptx`, `*.potx`, `*.ppt`,
  `*.pptm`, `*.odp` and `decks/`. Repo text (skills, docs, tests) names no company; the local
  test case is "a corporate master".
- **Clean room** (from `HANDOFF.md`): Anthropic's `pptx` skill is proprietary. It may be compared
  at the level of capabilities only; no code, text or lists are taken from it. Build from public
  sources: ECMA-376, python-pptx, LibreOffice and poppler documentation.
- **Local and offline at runtime.** No network calls; master and content never leave the machine.
- **Tooling:** Python via `uv`, `pyproject.toml` under `references/` (as in `diagram-skills`).
  Permissive dependencies only: python-pptx (MIT), lxml, Pillow, jsonschema. External tools
  invoked as processes, not bundled: LibreOffice (headless) and poppler (`pdftoppm`). PyMuPDF is
  excluded (AGPL).
- **Workspace rules** (`AGENTS.md`) apply, with the deviations listed in §11.
- **Every script:** `--help`, exit codes 0 = ok, 1 = check findings, 2 = usage or runtime error,
  and `--json` for machine-readable output.

## 4. How it is used

### 4.1 Start a deck

The user names a topic and a style source, for example "a deck on AI governance, style from
`~/Templates/company-master.potx` and `~/Templates/company-examples.pptx`".
The agent runs `init`, which:

1. creates `decks/<slug>/` in the user's project,
2. builds the working copy `decks/<slug>/<slug>.pptx`:
   - from a master: a new presentation on that master, no slides,
   - from an existing deck the user wants to continue: a copy with all slides,
   - from an existing deck used only as style: a copy with all slides removed (masters and
     layouts kept),
   - from a style file: the neutral template with the file's colors and fonts,
3. generates the style brief (§6), unless a cached one for the same sources exists.

### 4.2 Storyline (`deck-storyline` skill)

Input: free text, a topic, or files/folders the user points to (for example a research run).
The agent asks only what it needs (audience, live talk or pre-read, length, language), then
writes `decks/<slug>/storyline.md`:

- a header: purpose, audience, format, duration, language, style sources,
- one entry per slide: number, kicker, action title, what the slide shows (the visual idea in one
  or two sentences), source notes, status (`open` / `in design` / `approved` / `existing`), and
  once the slide exists, its stable slide id in the working copy (so entries survive inserts,
  moves and deletions).

When the deck continues an existing presentation (`--keep-slides`), the storyline starts with
one entry per existing slide (title and kicker read from the slide, status `existing`); new
entries are added around them.

Volker edits it in conversation or by hand. The storyline skill does not build slides.

### 4.3 Design one slide (`slide-design` skill)

For the next slide in the storyline, or any slide the user names:

1. **Read the state.** Load the style brief, read the storyline entry, and run `inventory` on the
   working copy. If the file changed since the agent last wrote it (hash differs), treat that as
   hand edits and work from the new state.
2. **Plan the slide** (in the conversation, briefly): the argument, the visual form that carries
   it (§8), the layout or example slide to start from, the grid.
3. **Apply** a change set (§7) with `apply`. A snapshot of the previous file is kept.
4. **Check and render** the changed slide with `check` and `render`.
5. **Review the PNG** against the visual checklist and fix what is wrong. At most three rounds of
   self-correction, then show the slide anyway and list what remains.
6. **Show** the slide: the PNG path (on macOS the agent may `open` it), a one-line summary of what
   it shows, and any open findings. Then wait.
7. **Revise** from feedback (text or screenshot) by going back to step 1. On approval, set the
   storyline status to `approved` and offer the next slide.

Slides can also be designed without a storyline (a single slide from a prompt); the skill then
skips the storyline steps.

### 4.4 Feedback by screenshot

The user sends a screenshot of a slide — from the agent's render or from PowerPoint, optionally
marked up — and says what bothers them. The agent:

1. identifies the slide (slide number in the footer, matching title text against the inventory,
   or asks if ambiguous),
2. maps the complaint to shapes in the inventory (by name, position and text),
3. changes those shapes with a change set and goes through check, render and review as above.

If the screenshot shows a slide that is not in the working copy, the user gives the deck it comes
from and the agent imports that slide (§4.5), or the agent builds a new native slide modeled on
the screenshot (structure from the screenshot, style from the style brief).

### 4.5 Getting slides back

The working copy is a normal `.pptx`. Volker opens it in PowerPoint and copies approved slides
into his own deck.

If he edits slides in his own deck and wants the agent to continue on them, he gives the agent
that file again. The agent **imports the edited slides into the working copy**, replacing their
older versions (the same cross-deck clone used for example slides, §7). It matches slides by
title against the inventory and asks when unsure. Slides in progress and the storyline stay
intact. A full re-`init` from his deck remains possible when he wants to start over from it.

If he edits the working copy itself in PowerPoint, the agent picks the changes up through step 1
of §4.3; he must save before asking the agent to continue and reload after the agent writes.
PowerPoint leaves a lock file (`~$<name>.pptx`) next to an open deck; when it is present the
agent says so before writing.

## 5. Architecture

### 5.1 Skills

| Skill | Job | Triggers on (examples) |
|---|---|---|
| `deck-storyline` | Turn text, a topic or research into `storyline.md`; refine it with the user | "build a storyline for …", "Storyline für ein Deck über …", "outline a presentation" |
| `slide-design` | Create, revise and check one slide in the working copy, in the style of a master | "design slide 3", "make this slide …", screenshot + "change …", "Folie zu …" |

Both document the boundary between them in a short table (pattern: `html-diagram` vs.
`excalidraw-diagram`). Both descriptions must be worded so they win over the built-in `pptx`
skill for these jobs; the user can always call them as `/powerpoint-skills:<skill>`.

### 5.2 Toolset

One CLI entry point, `pptx`, run as `uv run --project <plugin-root>/references pptx <command>`.

| Command | Does |
|---|---|
| `init <deck-dir> --from <source> [--examples <deck.pptx>] [--style-file <file>] [--keep-slides]` | Create the working copy and the style brief (§4.1) |
| `style <deck-dir> [--refresh]` | (Re)generate the style brief and contact sheets |
| `inventory <deck.pptx> [--slides N,…]` | List slides and shapes: stable id, name, kind, placeholder role, box in cm, text with run formatting, fill/line/font as theme references; also the way to read and extract content |
| `apply <deck.pptx> <changeset.json>` | Validate and apply a change set (§7), snapshot first, print what changed |
| `grid <deck-dir> --cols C [--rows R] [--gutter cm] [--area content\|full\|<box>]` | Compute aligned boxes inside the content area so the agent does not do grid arithmetic by hand |
| `render <deck.pptx> [--slides N,…]` | LibreOffice → PDF → PNG per slide (poppler), plus a contact sheet on request |
| `check <deck.pptx> [--slides N,…] [--all]` | Deterministic rules (§9), JSON report + text summary |
| `doctor` | Verify LibreOffice, `pdftoppm` and the fonts the style source needs; print install commands |

### 5.3 Files in the user's project

```
decks/<slug>/
  <slug>.pptx              # working copy — the agent's source of truth
  storyline.md             # deck-storyline output, status per slide
  deck.json                # style sources (paths), last written hash, settings
  style/
    brief.md               # style brief for the agent to read
    brief.json             # the same, machine-readable, used by apply/check
    layouts.png            # labeled contact sheet of all layouts
    examples.png           # labeled contact sheet of example slides (if any)
  changes/NNN-<what>.json  # every change set applied, in order
  snapshots/NNN.pptx       # file state before each change set
  render/slide-NN.png      # latest render per slide
  check/slide-NN.json      # latest check report per slide
```

Existing decks are never overwritten silently: `init` on an existing `decks/<slug>/` refuses
unless told to reuse or replace it.

### 5.4 Repo layout

```
powerpoint-skills/
  .claude-plugin/plugin.json
  .codex-plugin/plugin.json          # generated by scripts/sync-mirrors.sh
  skills/deck-storyline/SKILL.md
  skills/deck-storyline/references/storyline-format.md
  skills/slide-design/SKILL.md
  skills/slide-design/references/
    changeset-format.md              # ops, examples
    visual-forms.md                  # which form for which argument, with grid recipes
    review-checklist.md              # visual review checklist
  references/
    pyproject.toml, uv.lock
    pptx_skills/                     # the Python package behind the `pptx` CLI
    schemas/changeset.schema.json, brief.schema.json, style-file.schema.json
    tests/                           # pytest; fixtures generated at test time
  scripts/sync-mirrors.sh            # generates .codex-plugin/, --check for drift
  README.md, LICENSE
  docs/superpowers/specs/, docs/superpowers/plans/
```

## 6. Style sources and the style brief

`style` reads all given sources and writes `brief.md` (for the agent) and `brief.json` (for the
tools). The brief holds facts; it does not invent style.

- **Slide size.**
- **Theme:** color roles (`dk1`, `lt1`, `dk2`, `lt2`, `accent1`–`accent6`, links) with hex values,
  major and minor fonts.
- **Master text styles:** title size, body sizes per level. The minimum text size is **not**
  derived from them (real slides go below the master's body size); the default is 11 pt for
  body text and 9 pt for footnotes and sources, overridable in the style file.
- **Content area and grid:** from master or presentation guides if present, otherwise from the
  title and body placeholders of the content layouts; margins; footer zone.
- **Layouts:** name, placeholders with role (title, kicker/section label, body, picture, footer,
  number, date), box in cm, prompt text (to detect leftovers).
- **From example slides:** per slide the layout, shape count, fills and text colors used (as theme
  references with tints), font sizes used, typical gaps and corner radii, table styles. A summary
  of how colors are used (structure fills, emphasis, highlight) by frequency.
- **Contact sheets:** labeled PNGs of all layouts and all example slides.
- **Building blocks:** after the first `style` run the agent looks at the contact sheets once and
  adds a short catalog to `brief.md` ("example 36: lined table", "example 52: six KPI tiles").
  This section is agent-written and kept on refresh.
- **Rules:** defaults for the checks (§9), overridable by an optional rules section in the style
  file (for example max bullets per block, minimum font size).

The style file (for the "no PowerPoint source" case or to add rules to one) is JSON:
colors by theme role, major/minor font, optional rules. Its exact schema is left to the builder.

The **neutral template** used without a PowerPoint source is generated at runtime from
python-pptx's bundled default template, converted to 16:9, with the style file's colors and
fonts written into the theme. It is never committed as a binary.

## 7. Change sets

The agent never edits XML by hand and never regenerates a slide from scratch to change it. It
writes a change set (JSON, validated against `changeset.schema.json`) and `apply` performs it on
the current file. Change sets are kept in `changes/` so every step can be read and replayed.

Creating and revising are the same path: a new slide is a change set that starts with an add or
clone operation, a revision is one that changes existing shapes. There is no separate generation
engine.

Required operations in v1:

- **Slides:** add from a layout (by name), clone from an example slide or another slide in the
  working copy (with images, charts and relations; layouts mapped by name, clear error if a
  layout is missing), replace a slide with a slide from another deck (keeping its position and
  stable id), duplicate, delete, move.
- **Text:** set the text of a placeholder (by role) or shape (by id or name) as paragraphs of runs
  (bold, italic, theme color, size, bullet level, line breaks), keeping the existing formatting of
  anything not specified.
- **Shapes:** add rectangle, rounded rectangle, arrow/chevron forms, lines and connectors with
  arrowheads, text boxes; set position and size in cm; fill, line and text color as theme
  references with tint/shade; pattern fill (hatching); text insets and alignment; name every
  added shape meaningfully (`matrix-cell-r2-d3`).
- **Change existing shapes:** move, resize, restyle, re-text, delete, z-order, group/ungroup.
- **Tables:** add a native table, or fill a cloned example table; cell text and fills.
- **Charts:** native bar/column (clustered, stacked), line and doughnut charts with editable
  data, colored from the theme.
- **Pictures:** insert PNG/JPG into a picture placeholder or a box, keeping aspect ratio, cropped
  to fill; alt text.
- **Notes:** set speaker notes.

Raw RGB colors are rejected unless the op explicitly allows them; allowed ones are still
reported by `check`. Coordinates within 0.05 cm of a grid line snap to it. Exact op names and
field syntax are left to the builder.

## 8. Design rules

These go into `slide-design/SKILL.md` as the method (the counterpart to "diagrams should argue"
in `diagram-skills`). They are derived from the reference slide, the structure of the corporate
example deck and Volker's stated dislikes; marked (V) where Volker said it, otherwise my synthesis.

### Content

1. **One message per slide.** The action title states it as a full sentence (or two short ones),
   at most two lines at title size. If the title cannot be written, the slide is not ready.
2. **Kicker** = a short label for the slide's role or chapter, not a sentence.
3. **The body proves the title.** Whatever does not support the title goes to the notes or
   another slide.
4. **The structure carries the meaning.** Pick the visual form from the argument: matrix for two
   dimensions, process or chevrons for a sequence, comparison for alternatives, timeline for
   time, tiles for parallel facts, chart for numbers, table for lookup. `visual-forms.md` lists
   the forms with grid recipes.
5. **Little text, structured (V).** Labels plus one short explanation; bold lead-in followed by
   plain text; no text block longer than three lines; no bullet walls; detail goes into the
   speaker notes.
6. **Concrete over abstract.** Examples ("e.g. implement and test a requirement") instead of
   buzzwords.

### Visual

7. **Use the master.** Layouts and placeholders for kicker, title, footer and number; free shapes
   only inside the content area.
8. **Align to the grid.** Edges line up, gutters are equal, equal roles get equal sizes.
9. **Color has meaning.** Theme colors only. One color family for value/emphasis, neutral fills
   for structure, a highlight color at most once or twice per slide, the same meaning in the same
   color across the deck.
10. **Hierarchy through weight and size**, with few sizes per slide (the reference slide uses two
    body sizes plus the title), none below the brief's minimum.
11. **Native and editable.** No pictures of text, diagrams or tables.
12. **Consistent across slides.** Title position, kicker use and footer stay the same.

### Anti-patterns (each also a `check` warning where it can be detected)

- Keyword chains joined by " · ", " | " or " • " (V).
- Too much text: long paragraphs, bullet walls, shrinking the font to make it fit (V).
- Topic titles ("Overview", "Background") instead of statements.
- Decoration without meaning: icons, gradients, shadows, colors the master does not use.
- Everything in identical cards with no hierarchy.
- A "key takeaway" box repeating the title.
- Elements almost aligned; uneven gaps.
- Free text boxes where the layout has a placeholder.

## 9. Checks and review

### Deterministic (`check`)

Runs on the changed slides by default, so problems already in an existing deck do not drown the
report; `--all` checks the whole deck.

| Rule | Severity |
|---|---|
| File reopens with python-pptx; relationships resolve; no orphaned or duplicate parts; unique shape ids | error |
| Text overflows its shape (measured with the real font's metrics, wrapping at the box width minus insets) | error |
| Shape outside the slide | error |
| Leftover template text: layout prompt texts, "Lorem", "TODO", "XXX" | error |
| Font size below the brief's minimum (separate, lower minimum for footnotes and sources) | error |
| Shape outside the content area (except master elements) | warning |
| Overlapping text-bearing shapes that are not fully contained in one another | warning |
| Edges almost aligned (within 0.3 cm but not equal); unequal gutters in a row or column | warning |
| Colors not from the theme; fonts not from the theme | warning |
| Empty placeholders | warning |
| Text block longer than three lines; more than five bullets in a block | warning |
| Keyword chain (three or more items joined by " · ", " \| ", " • ") | warning |
| Title shorter than four words (likely a topic title) | info |
| Picture without alt text | info |
| A theme font is missing on this machine (renders will not match PowerPoint) | warning |

Thresholds come from the brief and can be overridden in the style file. They are calibrated
against the reference `.pptx`, which must pass with no errors and no text warnings (about 250
words in short structured blocks at 11–12 pt — so there is deliberately no words-per-slide
error). This is a local-only regression test (§12, call 10).

### Visual (the agent, every round)

`review-checklist.md`: the title states the takeaway; the eye follows the argument; hierarchy is
clear; edges and gaps are even; nothing is cramped at the edges; no awkward line breaks (single
word on the last line, broken compounds); readable contrast; consistent with the neighbouring
slides; the user's feedback is actually addressed. When the user's input was a screenshot,
compare the render with it side by side.

## 10. Multi-agent use

- Skills reference only shell commands and say "view the PNG with your image-viewing tool", never
  a Claude-specific tool name.
- Script paths are resolved from the skill's own directory (`<skill-dir>/../../references`).
- `scripts/sync-mirrors.sh` generates `.codex-plugin/plugin.json` from `.claude-plugin/plugin.json`
  with a skills pointer, and `--check` reports drift (pattern: `research-toolkit`).
- Acceptance includes running the slide workflow once in Codex CLI.

## 11. Changes against REQUIREMENTS.md and AGENTS.md

`REQUIREMENTS.md` and `HANDOFF.md` remain as background.

| Topic | REQUIREMENTS.md | v1 |
|---|---|---|
| Source of truth | YAML deck spec, `.pptx` rebuilt from it | The `.pptx`; change sets are applied to it (V) |
| Unit of work | Whole deck | Storyline, then one slide at a time (V) |
| Skill | One skill `pptx-deck` | Two skills: `deck-storyline`, `slide-design` |
| Screenshot | Rebuild foreign slides | Feedback on existing slides (V); rebuilding stays possible without extra tooling |
| Checks | Four stages incl. XSD and CI | Deterministic rules + visual review; no XSD, no CI |
| Style source | Master or neutral template | Master, existing deck, style file, or combinations (V) |

Deviations from the workspace `AGENTS.md`:

- **"One file under `references/` for style"** does not apply: style comes from the user's
  master; the plugin ships no brand style.
- **"Render with Playwright"**: rendering uses LibreOffice and poppler instead.
- **SKILL.md length:** each stays under 300 lines (stricter than the workspace's 500).

## 12. Decisions

Decided by Volker:

- Storyline first, then one slide at a time; v1 focuses on single slides and grows through use.
- Inputs: text, a topic, or prepared research.
- Feedback by text or by screenshot of an existing slide.
- Self-check and fix (overflow etc.) before showing, like `diagram-skills`.
- The `.pptx` is the source of truth; hand edits survive.
- The agent always works on its own local copy first.
- Works with any master, an existing deck as style source, or style specifications.
- Corporate masters and example slides are never committed.
- Usable by several agents; Anthropic's `pptx` skill as the model for capabilities.
- Dislikes: dot-joined keyword chains, too much text.

My calls (Volker may overturn them while reading):

1. Two skills, `deck-storyline` and `slide-design`; style setup is a step of `slide-design`, not a
   third skill.
2. Change sets as validated JSON applied by a tool, rather than the agent writing python-pptx
   code or raw XML.
3. Coordinates in cm; colors as theme roles with tint/shade.
4. LibreOffice + poppler for rendering; `brew install --cask libreoffice` and
   `brew install poppler` are setup steps in the README.
5. Getting slides back into the original is manual in v1.
6. The deck folder layout in §5.3, including snapshots and the change-set log.
7. Native charts are in v1 (python-pptx makes them cheap; research input often has numbers).
8. Check rules and thresholds in §9, calibrated against the reference slide.
9. The neutral template is generated at runtime, not committed.
10. Tests use generated fixtures. Local-only tests run against the corporate files when
    `PPTX_SKILLS_LOCAL_MASTER`, `PPTX_SKILLS_LOCAL_EXAMPLES` and `PPTX_SKILLS_LOCAL_REFERENCE`
    (the reference slide) point to them, and are skipped otherwise.
11. Edited slides from Volker's own deck are imported slide by slide into the working copy
    (§4.5) rather than replacing the whole working copy.
12. Storyline entries point to slides by stable slide id; continuing an existing deck lists its
    slides in the storyline as `existing`.

## 13. Left to the builder

- Exact change-set op names and field syntax; CLI argument details.
- The overflow algorithm (font lookup, line height, autofit handling).
- How placeholder roles (especially the kicker) are inferred across arbitrary masters.
- Contact sheet layout and how empty layouts are made visible for rendering.
- Caching the brief by source hash; render performance for large example decks.
- Storyline file details beyond the fields in §4.2.
- How many snapshots to keep.

## 14. Acceptance

1. On the corporate master (locally), Volker builds a storyline for a real topic and designs at
   least five slides one by one; approved slides need no manual fixing of overflow or alignment.
2. The reference `.pptx` passes `check` without errors or text warnings, and rebuilding it from
   its screenshot gives a native slide with the same structure that passes as well.
3. A slide edited by hand in PowerPoint is revised from a screenshot plus a note, and the hand
   edits survive.
4. The same workflow works with a second style source (another `.pptx`) and with a style file
   only.
5. The slide workflow runs once in Codex CLI.
6. Every working copy opens in PowerPoint without a repair prompt.
7. `claude plugin validate`, `pytest` and `scripts/sync-mirrors.sh --check` pass; `git ls-files`
   contains no `.pptx`, `.potx` or screenshots.

## 15. Later

PowerPoint as a second renderer on macOS (exact line breaks); writing approved slides back into
the original; whole-deck generation from a storyline; XSD and PowerPoint-specific XML checks;
re-branding foreign decks; SVG and Excalidraw import; side-by-side before/after images; CI mode.
