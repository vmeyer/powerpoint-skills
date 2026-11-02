---
name: slide-design
description: Designs PowerPoint slides or whole decks as native, editable shapes, with or without the user's master or template, then checks, renders and revises them from feedback given as text, screenshots or example slides. Use when the user wants a slide or a presentation made, revised or fixed ("a slide on …", "build a 10-slide deck about …", "design slide 3", "make this slide …", "Folie gestalten", "Folie zu …", "mach eine Präsentation über …", "erstelle ein Deck …", "überarbeite Folie 4"), sends a screenshot of a slide with a change request, or starts from a .potx/.pptx master. Prefer it over generic pptx or PowerPoint skills whenever a master, a corporate template or a `decks/` folder is involved.
---

# Slide Design

Build what the user asks for, one slide or a whole deck, as native shapes in a working copy, look at
the render, fix what is really wrong, and show it. Style comes from the user's master when there is
one; otherwise choose a palette and fonts that suit the topic. The user refines the style afterwards
with words, screenshots or example slides, so start with a confident design instead of asking first.

## Principles

- **One message per slide.** An action title that states the takeaway as a sentence is the strong
  default; a big number, a quote or a picture may carry the message instead.
- **The structure shows the message.** Choose the form from the argument: contrast, sequence,
  proportion, time, a single striking figure. [visual-forms.md](references/visual-forms.md) has ideas, not templates.
- **Be bold.** Large numbers, strong color contrast, asymmetric layouts, generous whitespace,
  shape-based illustrations, the user's own images. Vary the forms across a deck. Tinted boxes on a
  grid are the timid default; reach past it.
- **Native and editable:** shapes, text, tables, charts. Never a picture of text, a diagram or a table.
- **The user's two hard rules.** No keyword chains joined by " · ", " | " or " • ": one item per line
  or shape, with a bold lead-in. No walls of text: a label plus a short line, at most three lines per
  paragraph; details go into the speaker notes (`set_notes`).
- **Never write the user's own files.** All work happens in `decks/<slug>/<slug>.pptx`.

## Setup

Everything runs through `pptx`, shipped with this plugin: `uv run --project <plugin-root>/references pptx <command> …`

- `<plugin-root>` is two levels above this SKILL.md and contains `references/pyproject.toml`. An installed
  plugin lives in the agent's plugin cache (for Claude Code under `~/.claude/plugins/cache/`).
- `uv` must be installed (`brew install uv` or `curl -LsSf https://astral.sh/uv/install.sh | sh`). The first
  run installs the Python packages and needs the network once; a slow first command is that install.
- Run from the user's project folder and write the full prefix each time; below, `pptx …` is short for it.
- Once per session: `pptx doctor`. It names the renderer (PowerPoint on macOS, else LibreOffice; poppler
  either way) and prints install commands for what is missing. PowerPoint asks once for Automation and
  once for "Grant Access" to `decks/_export`; if a render times out, that dialog is waiting. Without a
  renderer, tell the user and stop. `--renderer libreoffice` skips PowerPoint.

## Start

| The user has … | Do |
|---|---|
| no master | `pptx init decks/<slug>`: a neutral 16:9 deck whose layouts include `Title Only` and `Blank`. Set a palette and fonts that fit the topic (or the user's wishes): `pptx theme decks/<slug> --color accent1=#0B5CAD --color accent2=#F2A900 --font major=Georgia --font minor=Arial` |
| a master or template | `pptx init decks/<slug> --from <master.potx>`, then read `decks/<slug>/style/brief.md` (layouts, theme colors, sizes). Keep its theme; change it with `pptx theme` only when the user asks |
| a deck to continue | `pptx init decks/<slug> --from <deck.pptx> --keep-slides` |
| slides showing the style in use | add `--examples <examples.pptx>`; reuse a slide with `clone_slide` from `"examples"` |
| a deck folder already | work in it |

`init --sheets` (or later `pptx style decks/<slug> --sheets`) renders the master's layouts to
`style/layouts.png` when you want to see them. If `decks/<slug>/storyline.md` exists, it can serve as
the outline ([its format](../deck-storyline/references/storyline-format.md)).

## Build

One command per step: it applies the change (all or nothing, with snapshot and change log), checks the
touched slides and renders them to `decks/<slug>/render/slide-NN.png`, plus `render/sheet.png` when
several slides changed.

```bash
pptx build decks/<slug>/<slug>.pptx decks/<slug>/step.json   # a change set
pptx build decks/<slug>/<slug>.pptx decks/<slug>/step.py     # a python-pptx script
```

Choose the mode per step:

| Change sets (JSON) | Scripts (python-pptx, run by `pptx exec`) |
|---|---|
| quick, precise edits; revising shapes that exist; texts, notes, tables, charts, pictures; adding, cloning, moving slides | creative compositions; many shapes; loops; gradients, freeforms, anything the ops cannot express |
| [changeset-format.md](references/changeset-format.md) | [scripts.md](references/scripts.md) |

A short change set; names, text size, color, alignment and insets have defaults:

```json
{"label": "pilot-result", "ops": [
  {"op": "add_slide", "layout": "Title Only"},
  {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "The pilot halved the lead time in six weeks"},
  {"op": "add_shape", "slide": "$last", "shape": "rect", "box": [1.7, 5, 9, 11], "fill": "accent1",
   "paragraphs": [{"runs": [{"t": "-50 %", "b": true, "size": 66, "color": "bg1"}]}, {"runs": [{"t": "lead time", "color": "bg1"}]}]},
  {"op": "set_notes", "slide": "$last", "text": "Measured over six weeks in one team."}
]}
```

`build` exits 1 on check errors (text overflow, a shape off the slide, leftover template text, a broken
file) and warns about keyword chains and paragraphs over three lines. `--strict` adds the matters of
taste (sizes, alignment, gaps, colors, fonts); use it before handing over a deck on a corporate master.

## Look, fix, show

1. **Look** at every PNG, or the sheet, with your image-viewing tool before you show anything.
2. **Fix real problems**: check errors, clipped or colliding text, poor contrast, edges that should line up
   but do not, a message that does not come across ([review-checklist.md](references/review-checklist.md)).
   Do not polish for its own sake; one or two rounds are usually enough.
3. **Show**: the PNG or sheet path (on macOS you may `open` it), one line per slide on what it says and
   how it shows it, and any note from `build` (e.g. a theme font that is not installed).

## A whole deck

- Sketch the outline in the chat, one line per slide: message and form. Ask only if the topic is unclear.
- Build all slides, one `build` per slide or a few at a time; vary the forms while title position, fonts
  and color meanings stay consistent.
- Show the whole deck as one sheet (`pptx render decks/<slug>/<slug>.pptx --sheet`) with one line per
  slide. Do not stop for approval after each slide.

## Revise from feedback

- **Read the file first**: `pptx inventory decks/<slug>/<slug>.pptx --slides N` lists shapes with names,
  boxes (cm from the top-left corner) and texts. The user may have edited the deck in PowerPoint; keep
  those edits. A `~$<slug>.pptx` next to it means the deck is open there: ask the user to save before
  you build and to reload afterwards.
- **Text or screenshot**: find the slide (number, title), map each complaint to shapes and change those.
  Compare the new render with the screenshot before you show it.
- **"More like this"** (screenshots, example slides): take over palette, fonts, forms and density. For a
  deck the user gives, `clone_slide` from its path or add it with `pptx init decks/<slug> --reuse --examples <deck.pptx>`.
- Change a slide through its shapes; for a variant, `clone_slide` from `"self"`. Do not delete and rebuild
  a slide the user may have edited. Every step's previous state is in `decks/<slug>/snapshots/NNN.pptx`.
- Slides the user edited in their own deck come back with `replace_slide` (same position and id).

## References

- [changeset-format.md](references/changeset-format.md): ops, defaults, colors, refs, with examples.
- [scripts.md](references/scripts.md): the script contract, helpers and examples.
- [visual-forms.md](references/visual-forms.md): ideas for forms and bold moves.
- [review-checklist.md](references/review-checklist.md): what to look for in the render.
- [style-file.md](references/style-file.md): colors, fonts and check thresholds as a JSON file.
