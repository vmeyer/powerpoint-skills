# powerpoint-skills

A plugin for Claude Code, Codex CLI and GitHub Copilot CLI that designs PowerPoint slides, one slide or a whole deck, as native, editable shapes, with or without your master. Every step is checked and rendered to PNG before you see it, and you refine the result with words, screenshots or example slides.

The [v1 spec](docs/superpowers/specs/2026-11-02-powerpoint-skills-v1-design.md) holds the original design; `REQUIREMENTS.md` and `HANDOFF.md` are German background notes.

## Skills

Both skills are also available as `/powerpoint-skills:<skill>`.

### slide-design
Builds what you ask for, a single slide or a complete deck, in a working copy of your master or on a neutral 16:9 template with a palette and fonts that suit the topic. It renders every step, looks at the result, fixes real problems and shows you the PNG or a contact sheet. It also revises slides from feedback as text, a screenshot or example slides.

> "Make a slide on why small batches cut lead time."
> "Build a 6-slide deck about our pilot results, dark blue and orange."

### deck-storyline
Optional planning before a longer deck: one entry per slide (kicker, action title, what the slide shows, sources, status) in `decks/<slug>/storyline.md`. `slide-design` can use it as the outline.

> "Build a storyline for a 10-slide deck from this research."

| Use `slide-design` when … | Use `deck-storyline` when … |
|---|---|
| a slide or a deck is created, revised or fixed | the sequence of messages for a long deck is to be agreed first |
| you give feedback as text, screenshot or example slides | you bring text, a topic or research and want an outline |
| result: slides in `decks/<slug>/<slug>.pptx` | result: `decks/<slug>/storyline.md` |

## How it works

1. **Start.** `pptx init decks/<slug>` creates the working copy: on your master (`--from master.potx`) or, without one, on a neutral 16:9 template whose colors and fonts `pptx theme` sets.
2. **Build.** Each step is one `pptx build` with either a short JSON change set (texts, shapes, tables, charts, pictures, notes, slide operations) or a python-pptx script for free compositions (`pptx exec` runs it with the same safety). `build` applies the step all or nothing, checks the touched slides and renders them, with a contact sheet when several slides changed.
3. **Look and show.** The agent views the PNGs, fixes what is really wrong (overflowing text, shapes off the slide, leftover template text, poor contrast) and shows you the result. `--strict` adds the matters of taste: sizes, alignment, gaps, colors, fonts.

Every change set or script is logged and the file before it kept as a snapshot, so each step can be traced and undone. Your own files are never written.

## Install and use in Claude Code, Codex CLI and GitHub Copilot CLI

The plugin is not in the `vmeyer-plugins` marketplace yet. Until its first release, install it from this
repository as shown under "Before the release"; the marketplace lines work once it is listed. Each agent then
needs the engine set up once (see [Setup](#setup)).

| | Claude Code | Codex CLI | GitHub Copilot CLI |
|---|---|---|---|
| Reads | `.claude-plugin/plugin.json` | `.codex-plugin/plugin.json` | `.claude-plugin/plugin.json` |
| Uses a skill | when the request matches it, or `/powerpoint-skills:<skill>` | when the request matches it, or a `$` mention | when the request matches it |

`.codex-plugin/plugin.json` is generated from `.claude-plugin/plugin.json` by `scripts/sync-mirrors.sh`
and points Codex at `skills/`. Copilot CLI reads `.claude-plugin/plugin.json` directly and ignores the Codex file.

### Claude Code

```
/plugin marketplace add vmeyer/claude-plugins
/plugin install powerpoint-skills@vmeyer-plugins
```

The same from a shell: `claude plugin marketplace add vmeyer/claude-plugins`, then
`claude plugin install powerpoint-skills@vmeyer-plugins`.

Before the release, or for development, load a clone for one session:

```bash
git clone https://github.com/vmeyer/powerpoint-skills.git
claude --plugin-dir ./powerpoint-skills
```

Ask for what you want ("Make a slide on …", "Build a deck about …") or call a skill by name:
`/powerpoint-skills:deck-storyline`, `/powerpoint-skills:slide-design`.

### Codex CLI

Codex installs plugins from marketplaces and reads the same `marketplace.json` as Claude Code:

```bash
codex plugin marketplace add vmeyer/claude-plugins
codex plugin add powerpoint-skills@vmeyer-plugins
```

Before the release, list a clone in a local marketplace: a folder `local-plugins/` that contains the clone
as `powerpoint-skills/` and a file `.claude-plugin/marketplace.json`:

```json
{
  "name": "local-plugins",
  "owner": { "name": "you" },
  "plugins": [{ "name": "powerpoint-skills", "source": "./powerpoint-skills" }]
}
```

```bash
codex plugin marketplace add ./local-plugins
codex plugin add powerpoint-skills@local-plugins
```

Codex copies the plugin folder into its cache (`~/.codex/plugins/cache/<marketplace>/powerpoint-skills/<version>/`),
untracked files included. Install from a fresh clone, never from a folder that also holds masters or decks.

Codex uses a skill when your request matches its description. To name one, mention it with `$`, e.g.
`$slide-design`; the mention list shows the exact names.

### GitHub Copilot CLI

```bash
copilot plugin marketplace add vmeyer/claude-plugins
copilot plugin install powerpoint-skills@vmeyer-plugins
```

Before the release, install straight from GitHub or load a clone for one session:

```bash
copilot plugin install vmeyer/powerpoint-skills   # Copilot marks direct installs as deprecated
copilot --plugin-dir ./powerpoint-skills
```

The local marketplace from the Codex section works too (`copilot plugin marketplace add ./local-plugins`, then
`copilot plugin install powerpoint-skills@local-plugins`); Copilot loads the clone in place instead of copying it.

`copilot plugin list` shows the plugin and `copilot skill list` (in a session: `/skills`) shows
`deck-storyline` and `slide-design`. Copilot uses a skill when your request matches its description.

## Setup

Set the engine up once per machine, or ask your agent: "Set up the powerpoint-skills renderer."

1. **uv** runs the Python engine. macOS: `brew install uv`; macOS or Linux without Homebrew, the official
   installer: `curl -LsSf https://astral.sh/uv/install.sh | sh`.
2. **A renderer** turns the deck into a PDF. The engine uses **Microsoft PowerPoint** when it is installed
   (macOS, in `/Applications` or `~/Applications`) and **LibreOffice** otherwise. With PowerPoint there is
   nothing to install, but it needs a one-time permission, given on the first render:

   - **Automation:** macOS asks whether your terminal (or the app your agent runs in) may control
     Microsoft PowerPoint. Allow it. The setting lives under System Settings → Privacy & Security → Automation.
   - **Folder access:** PowerPoint is sandboxed and writes only into folders granted one by one (granting a
     parent folder does not count). The engine therefore hands every deck of a project to PowerPoint through
     one fixed folder, `decks/_export` (for a `.pptx` outside a deck folder: `_export` next to it). The first
     PowerPoint export in a project shows one "Grant Access" dialog for `decks/_export`: click
     **Grant Access** once. If you click **Cancel**, the render falls back to LibreOffice where it is installed.

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

   To pick a renderer yourself, pass `--renderer powerpoint` or `--renderer libreoffice` to `pptx build`,
   `pptx render`, `pptx init` and `pptx style`, or set `PPTX_SKILLS_RENDERER` (default `auto`).
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
plugin cache, `~/.claude/plugins/cache/<marketplace>/powerpoint-skills/<version>/` (Codex:
`~/.codex/plugins/cache/<marketplace>/powerpoint-skills/<version>/`). The first `uv run` installs
the engine's Python packages (python-pptx, lxml, Pillow, jsonschema) into `references/.venv`; that needs the
network once, and every later run works offline.

`pptx doctor` names the renderer in use, shows PowerPoint's version and one-time setup, and reports what is
missing (a renderer, poppler, fonts used by your master).

Every command reports a problem as `error: <message>` and exits with 2 (1 means `check` or `build` found errors).
For an unexpected error, set `PPTX_SKILLS_DEBUG=1` to also print the Python traceback.

## Customization

Slides follow your style when you have one and a fitting one when you do not. Style can be refined at any
time: describe it, send screenshots or point to example slides ("make it more like this").

### Without a master

`pptx init decks/<slug>` starts on a neutral 16:9 template. The agent picks colors and fonts that suit the
topic, or takes yours:

```bash
pptx theme decks/<slug> --color accent1=#0B5CAD --color accent2=#F2A900 --font major=Georgia --font minor=Arial
```

`pptx theme decks/<slug>` without options shows the current theme. A [style file](skills/slide-design/references/style-file.md)
sets colors, fonts and check thresholds at `init` when several decks share them.

### Where to keep masters and examples

Masters, templates and example decks are usually confidential. Keep them outside any git repository, for
example in `~/Documents/pptx-templates/`, and never put them into a public repository. `init` stores the
absolute path of every source in `decks/<slug>/deck.json` and reads them from there, so keep them at a fixed
place; after a move, point the deck at the new paths with `pptx init decks/<slug> --reuse --from … --examples …`.
The working copy `decks/<slug>/<slug>.pptx` is a copy of your master: if your project repository is public or
shared, add `decks/` and `_export/` (PowerPoint's exchange folder, which may briefly hold a copy of a deck) to its
`.gitignore`.

### Style sources

| Source | Option | What `init` does with it |
|---|---|---|
| Master or template (`.potx`, or a `.pptx` used as one) | `--from` | Copies it to the working copy (its slides are dropped unless `--keep-slides`, which continues an existing deck) and reads layouts, placeholder roles, theme colors and fonts, text sizes, guides and table styles into `style/brief.md`. With a master, the agent changes the theme only when you ask. |
| Examples deck (`.pptx`) | `--examples` | Reads it in place, never copies it. The brief records per slide the layout, colors, font sizes and table styles; single slides can be cloned into the deck as building blocks. |
| Style file (JSON) | `--style-file` | Colors and fonts for the neutral template, and `rules` that override check thresholds. |
| nothing | | The neutral 16:9 template; `pptx theme` sets its colors and fonts. |

`init --sheets` (or `pptx style decks/<slug> --sheets`) also renders contact sheets of the layouts and examples
to `style/layouts.png` and `style/examples.png`; without it `init` skips them and is about 30 s faster.

### What makes a master work well

- **Guides define the content area.** Put exactly two vertical and two horizontal guides on the slide master
  (Slide Master view → Guides). The area between them is where content goes, and the checks measure
  alignment against it. Without such guides the area is taken from the layouts' placeholders.
- **A small kicker placeholder above the title.** A body placeholder at most 1.5 cm high that sits above the
  title is recognized as the kicker, the short label over each action title.
- **Meaningful layout names.** Slides are added by layout name, and the brief lists layouts by name:
  "Title only" or "Two columns" say what they are for, "Layout 7" does not.
- **A title-only layout** (kicker and title, nothing else). Most designed slides start from it.
- **Footer, slide number and date placeholders on the layouts.** They are copied onto every new slide.
- **Table styles in the master.** Define your table styles in the template and set the default one there:
  new tables get the default, others can be chosen by name.
- **Theme colors and fonts in the theme**, not hard-coded on shapes. The brief reads the theme, and change
  sets and scripts refer to theme colors such as `accent1`, so slides follow the master. Install the theme fonts
  on the machine; `pptx doctor decks/<slug>` checks them.

### What makes a good examples deck

- **One slide per building block:** tables, KPI tiles, a process, a matrix, a chart, a picture with text.
- **Real-looking but neutral text**, without client names or confidential figures: the slides may appear in
  `style/examples.png` and be cloned into new decks.
- **Built on the same master**, so layout names match.
- **No hidden slides and no embedded OLE objects** (e.g. embedded Excel sheets): a slide with an OLE object
  cannot be cloned.

## Output Convention

Everything for a deck lives in the project of the user, under `decks/<slug>/` with the slug in kebab-case:

```
decks/<slug>/
  <slug>.pptx                 # working copy, the source of truth
  storyline.md                # deck-storyline output (optional)
  deck.json                   # style sources, last written hash, settings
  style/                      # style brief (md + json); contact sheets with --sheets
  changes/NNN-<what>.json|.py # every change set, script or theme change, in order
  snapshots/NNN.pptx          # file state before each change
  render/slide-NN.png         # latest render per slide
  render/sheet.png            # contact sheet of the last build or render --sheet
  check/slide-NN.json         # latest check report per slide
decks/_export/                # PowerPoint's exchange folder, shared by all decks; empty between renders
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
      SKILL.md                        # Principles and workflow for slides and decks
      references/
        changeset-format.md           # Change set operations, defaults and examples
        scripts.md                    # Script contract for pptx exec, with examples
        visual-forms.md               # Ideas for forms and bold moves
        review-checklist.md           # What to look for in a render
        style-file.md                 # Style file format and check thresholds
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
