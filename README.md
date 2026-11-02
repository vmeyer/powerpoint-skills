# powerpoint-skills

A plugin for Claude Code, Codex CLI and GitHub Copilot CLI that designs PowerPoint slides **one at a time** in the style of any master or deck. A storyline comes first, then each slide is built from native, editable shapes in a working copy, rendered to PNG and checked before you see it.

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

Ask for what you want ("Design slide 3 of the storyline") or call a skill by name:
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
plugin cache, `~/.claude/plugins/cache/<marketplace>/powerpoint-skills/<version>/` (Codex:
`~/.codex/plugins/cache/<marketplace>/powerpoint-skills/<version>/`). The first `uv run` installs
the engine's Python packages (python-pptx, lxml, Pillow, jsonschema) into `references/.venv`; that needs the
network once, and every later run works offline.

`pptx doctor` names the renderer in use, shows PowerPoint's version and one-time setup, and reports what is
missing (a renderer, poppler, fonts used by your master).

Every command reports a problem as `error: <message>` and exits with 2 (1 means `check` found errors). For an
unexpected error, set `PPTX_SKILLS_DEBUG=1` to also print the Python traceback.

## Customization

Slides follow your style. This section says what to prepare, where to keep it and what makes it work well.

### Where to keep masters and examples

Masters, templates and example decks are usually confidential. Keep them outside any git repository, for
example in `~/Documents/pptx-templates/`, and never put them into a public repository. `init` stores the
absolute path of every source in `decks/<slug>/deck.json` and reads them from there, so keep them at a fixed
place; after a move, point the deck at the new paths with `pptx init decks/<slug> --reuse --from … --examples …`.
The working copy `decks/<slug>/<slug>.pptx` is a copy of your master: if your project repository is public,
add `decks/` to its `.gitignore`.

### Three style sources

| Source | Option | What `init` does with it |
|---|---|---|
| Master or template (`.potx`, or a `.pptx` used as one) | `--from` | Copies it to the working copy (its slides are dropped unless `--keep-slides`, which continues an existing deck) and reads layouts, placeholder roles, theme colors and fonts, text sizes, guides and table styles into `style/brief.md`. `style/layouts.png` shows every layout. |
| Examples deck (`.pptx`) | `--examples` | Reads it in place, never copies it. The brief records per slide the layout, colors, font sizes and table styles, plus typical gaps and corner radii; `style/examples.png` shows the slides. Single slides can be cloned into the deck as building blocks. |
| Style file (JSON) | `--style-file` | Without a master: colors and fonts for a neutral 16:9 template. With or without one: `rules` that override check thresholds (minimum font size, maximum bullets, snap distance, …). See [`style-file.md`](skills/slide-design/references/style-file.md). |

`init` always needs `--from` or `--style-file`; `--examples` and a rules-only style file can be added to either.

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
  sets refer to theme colors such as `accent1`, so slides follow the master. Install the theme fonts on the
  machine; `pptx doctor decks/<slug>` checks them.

### What makes a good examples deck

- **One slide per building block:** tables, KPI tiles, a process, a matrix, a chart, a picture with text.
- **Real-looking but neutral text**, without client names or confidential figures: the slides appear in
  `style/examples.png` and may be cloned into new decks.
- **Built on the same master**, so layout names match.
- **No hidden slides and no embedded OLE objects** (e.g. embedded Excel sheets): a slide with an OLE object
  cannot be cloned.

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
