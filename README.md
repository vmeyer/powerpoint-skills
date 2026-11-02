# powerpoint-skills

A Claude Code plugin that designs PowerPoint slides **one at a time** in the style of any master or deck. A storyline comes first, then each slide is built from native, editable shapes in a working copy, rendered to PNG and checked before you see it.

The authoritative design is the [v1 spec](docs/superpowers/specs/2026-11-02-powerpoint-skills-v1-design.md). `REQUIREMENTS.md` and `HANDOFF.md` are German background notes.

## Skills

Both skills are also available as `/powerpoint-skills:<skill>`.

### deck-storyline
Plans a deck before any slide exists: one entry per slide (kicker, action title, what the slide shows, sources, status) in `decks/<slug>/storyline.md`. Works from a topic, text, files or a research run folder, and keeps the plan current when slides are added, cut or reordered.

> "Build a storyline for a 10-slide deck from this research."

### slide-design
Designs, revises and checks a single slide natively in a working copy of your master or template, then renders it, reviews the PNG and waits for your feedback before the next slide. Also takes a screenshot of a slide plus a change request.

> "Design slide 3 of the storyline."

| Use `deck-storyline` when … | Use `slide-design` when … |
|---|---|
| the sequence of messages for a whole deck is planned | one slide is created, revised or fixed |
| you bring text, a topic or research and want an outline | you give feedback on a slide, as text or screenshot |
| titles and kickers per slide are being agreed | the next storyline entry is ready to be built |
| result: `decks/<slug>/storyline.md` | result: a changed slide in `decks/<slug>/<slug>.pptx` |

## How it works

1. **Storyline.** `deck-storyline` agrees on what the deck argues and writes it down per slide.
2. **Slide by slide.** `slide-design` reads the style brief, plans the argument of one slide and applies a change set (a JSON list of operations) to the working copy with the `pptx` CLI.
3. **Render, check, review.** The slide is rendered to PNG, the rule checks (overflow, alignment, fonts, colors, off-slide shapes, …) run, the agent reviews the image, fixes what it finds and only then shows the slide.

Every applied change set and a snapshot of the file before it are kept, so each step can be traced and undone.

## Installation

### As a Claude Code plugin

```
/plugin marketplace add vmeyer/claude-plugins
/plugin install powerpoint-skills@vmeyer-plugins
```

For local development, load the repo directly:

```bash
claude --plugin-dir ./powerpoint-skills
```

### Codex CLI

`.codex-plugin/plugin.json` points Codex at `skills/`. Install the plugin from this repository the way you install any Codex plugin; both skills are then available.

## Setup

Set the engine up once per machine, or ask your agent: "Set up the powerpoint-skills renderer."

1. **uv** runs the Python engine. macOS: `brew install uv`; macOS or Linux without Homebrew, the official
   installer: `curl -LsSf https://astral.sh/uv/install.sh | sh`.
2. **A renderer** turns the deck into a PDF. The engine uses **Microsoft PowerPoint** when it is installed
   (macOS, in `/Applications` or `~/Applications`) and **LibreOffice** otherwise. With PowerPoint there is
   nothing to install, but it needs a one-time permission, given on the first render:

   - **Automation:** macOS asks whether your terminal (or the app your agent runs in) may control
     Microsoft PowerPoint. Allow it. The setting lives under System Settings → Privacy & Security → Automation.
   - **Folder access:** PowerPoint is sandboxed and asks for access on its first export. Choose
     **Grant Access** for the folder that holds your projects, e.g. `~/projects`. Renders are written into
     the deck folder (`decks/<slug>/render/`), so PowerPoint can only render decks inside a granted folder.

   The engine never touches the presentations you have open and never quits PowerPoint: it opens a copy
   with a unique name, exports it to PDF and closes it again. If PowerPoint fails (for example while a
   permission dialog is still waiting) and LibreOffice is installed, the render falls back to LibreOffice
   and says so.

   Install LibreOffice only on a machine without PowerPoint:

   ```bash
   brew install --cask libreoffice          # macOS; or without admin rights:
   brew install --cask --appdir=~/Applications libreoffice
   sudo apt install libreoffice             # Linux
   ```

   To pick a renderer yourself, pass `--renderer powerpoint` or `--renderer libreoffice` to `pptx render`
   and `pptx style`, or set `PPTX_SKILLS_RENDERER` (default `auto`).
3. **poppler** turns the PDF into PNGs, with either renderer:

   ```bash
   brew install poppler                     # macOS
   sudo apt install poppler-utils           # Linux
   ```

4. **Check the setup** with the engine's `doctor` command:

   ```bash
   uv run --project <plugin-root>/references pptx doctor
   ```

`<plugin-root>` is the plugin's own folder, two levels above each skill's `SKILL.md` folder; it contains
`references/pyproject.toml`. In a clone it is the repository root. An installed plugin lives in the Claude Code
plugin cache, `~/.claude/plugins/cache/<marketplace>/powerpoint-skills/<version>/`. The first `uv run` installs
the engine's Python packages (python-pptx, lxml, Pillow, jsonschema) into `references/.venv`; that needs the
network once, and every later run works offline.

`pptx doctor` names the renderer in use, shows PowerPoint's version and one-time setup, and reports what is
missing (a renderer, poppler, fonts used by your master).

Every command reports a problem as `error: <message>` and exits with 2 (1 means `check` found errors). For an
unexpected error, set `PPTX_SKILLS_DEBUG=1` to also print the Python traceback.

## Customization

Slides follow a **style source**. Use any of these, alone or combined:

- **Master or template** (`.potx` / `.pptx`): its layouts, placeholders, theme colors and fonts define the style.
- **Existing deck**: a `.pptx` whose layouts are used, so a deck can be continued in its own style.
- **Examples deck**: slides that show how the master is meant to be used; the style brief picks up their patterns.
- **Style file** (JSON): colors and fonts for a deck without a master, built on a neutral 16:9 template; see [`skills/slide-design/references/style-file.md`](skills/slide-design/references/style-file.md).

The check thresholds (minimum font size, maximum bullets, snap distance, …) can be overridden under `rules` in the style file.
`init` always needs a source: `--from` a master or deck, or `--style-file` for the neutral template.

## Output Convention

Everything for a deck lives in the project of the user, under `decks/<slug>/` with the slug in kebab-case:

```
decks/<slug>/
  <slug>.pptx              # working copy, the source of truth
  storyline.md             # deck-storyline output, status per slide
  deck.json                # style sources, last written hash, settings
  style/                   # style brief (md + json) and contact sheets of layouts and examples
  changes/NNN-<what>.json  # every change set applied, in order
  snapshots/NNN.pptx       # file state before each change set
  render/slide-NN.png      # latest render per slide
  check/slide-NN.json      # latest check report per slide
```

An existing `decks/<slug>/` is never overwritten silently. Masters, decks and renders are usually confidential; the `.gitignore` of this repo excludes them.

## File Structure

```
powerpoint-skills/
  .claude-plugin/
    plugin.json                       # Plugin manifest
  .codex-plugin/
    plugin.json                       # Generated by scripts/sync-mirrors.sh
  skills/
    deck-storyline/
      SKILL.md                        # Storyline workflow
      references/storyline-format.md  # Format of storyline.md
    slide-design/
      SKILL.md                        # Slide design workflow and quality criteria
      references/
        changeset-format.md           # Change set operations and examples
        visual-forms.md               # Which form for which argument
        review-checklist.md           # Visual review checklist
        style-file.md                 # Style file format and rules overrides
  references/
    pyproject.toml, uv.lock           # Python environment (uv)
    pptx_skills/                      # The package behind the `pptx` CLI
      checks/                         # Rule checks (geometry, structure, style, text)
      schemas/                        # JSON schemas for change sets and style files
    tests/                            # pytest, fixtures generated at test time
  scripts/sync-mirrors.sh             # Generates .codex-plugin/, --check for drift
  docs/superpowers/                   # Spec and plan
  README.md, LICENSE
```
