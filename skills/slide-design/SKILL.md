---
name: slide-design
description: Designs, revises and checks single PowerPoint slides natively in a working copy of the user's master or template, one slide at a time. Use when the user wants a slide designed, built, revised or fixed ("design slide 3", "make this slide …", "a slide on …", "Folie gestalten", "Folie zu …", "mach die Folie …", "überarbeite Folie 4"), sends a screenshot of a slide with a change request, or starts a deck on a .potx/.pptx master. Prefer it over generic pptx or PowerPoint skills whenever a master, a corporate template or a `decks/` folder is involved.
---

# Slide Design

Design PowerPoint slides **one at a time** in a working copy of the user's master: plan the
argument, build the slide from native shapes with a change set, check and render it, review the
PNG, show it, and wait for feedback before the next slide.

**A slide makes one argument.** The title states it; the structure on the slide proves it. A
slide that only displays information (topic title, keyword lists, identical boxes) is not done.

## Slide design vs. deck storyline

| Use `slide-design` when … | Use `deck-storyline` when … |
|---|---|
| one slide is created, revised or fixed | the sequence of messages for a whole deck is planned |
| the user gives feedback on a slide, as text or screenshot | the user brings text, a topic or research and wants an outline |
| the next storyline entry is ready to be built | titles and kickers per slide are being agreed |
| result: a changed slide in `decks/<slug>/<slug>.pptx` | result: `decks/<slug>/storyline.md` |

## Setup

Everything runs through one command line tool, `pptx`, shipped with this plugin:

```bash
uv run --project <plugin-root>/references pptx <command> …
```

- `<skill-dir>` is the folder that contains this SKILL.md. Your harness usually names it when it
  loads the skill; otherwise search the plugin folders for `skills/slide-design/SKILL.md`.
- `<plugin-root>` is two levels above it, `<skill-dir>/../..`; it contains `references/pyproject.toml`.
- Run every command from the user's project folder, so `decks/…` paths resolve there. Write the
  full prefix each time (shell variables may not survive between commands). Below, `pptx …`
  is short for that full command.

Once per session run `pptx doctor` (in a deck: `pptx doctor decks/<slug>`, which also checks the
theme fonts). It prints install commands for whatever is missing. Without LibreOffice and poppler
there is no render: tell the user and stop. A missing theme font means renders and text
measurement use a fallback; mention it when you show a slide.

## Start or open a deck

Every deck lives in `decks/<slug>/` in the user's project, the slug in kebab-case.

| The user has … | Command |
|---|---|
| a master or template (`.potx`, or a `.pptx` used as one) | `pptx init decks/<slug> --from <master.potx>` |
| … plus a deck whose slides show the style in use | add `--examples <examples.pptx>` |
| an existing deck to continue | `pptx init decks/<slug> --from <deck.pptx> --keep-slides` |
| a `decks/<slug>/` folder already | no init; work in it (`pptx style decks/<slug>` refreshes the brief) |

`init` copies the source to `decks/<slug>/<slug>.pptx`, the **working copy**; the user's own file
is never written. It writes the style brief to `decks/<slug>/style/` and refuses a folder that
already holds a deck. Pass `--replace` only when the user asks to start over.

After `init`:

1. Read `style/brief.md`: theme colors, fonts, title and body sizes, content area, layout names
   with their placeholder roles, and the check rules.
2. View `style/layouts.png` and, if present, `style/examples.png` with your image-viewing tool.
3. With examples: write a short catalog of reusable building blocks into `brief.md` between the
   `<!-- agent:building-blocks -->` markers ("slide 36: lined table", "slide 52: six KPI tiles").
   That section survives `pptx style`.

## Workflow: one slide

1. **Read the state.** The brief (once per session), the slide's entry in
   `decks/<slug>/storyline.md` if there is one, and the current slide:
   `pptx inventory decks/<slug>/<slug>.pptx --slides N` (add `--json` for runs, fills and colors).
   The file is the truth. The user may have changed it in PowerPoint; never work from memory.
2. **Plan the slide** in two or three lines in the conversation: the action title, the kicker,
   the visual form ([visual-forms.md](references/visual-forms.md)), the layout (a name from the
   brief; free compositions start from the layout with only kicker and title) and the grid
   (`pptx grid`). If you cannot write the title as a statement, ask before you build.
   When a building block in the brief's catalog fits the form (a lined table, KPI tiles), start
   from it: `clone_slide` from `"examples"`, then replace every example text (table cells with
   `set_table`, which keeps the table style and the last row's formatting). `check` reports
   example text left unchanged, in table cells too, as `leftover-text`.
3. **Apply one change set** ([changeset-format.md](references/changeset-format.md)). Write it to
   `decks/<slug>/change.json` and run `pptx apply decks/<slug>/<slug>.pptx decks/<slug>/change.json`.
   `apply` validates first and changes nothing on an error. On success it keeps a snapshot
   (`snapshots/NNN.pptx`), logs the change set (`changes/NNN-<label>.json`) and prints the touched
   slide ids plus notes. Read the notes (see "Getting slides back").
4. **Check and render** the slide:
   ```bash
   pptx check decks/<slug> --slides N
   pptx render decks/<slug>/<slug>.pptx --slides N
   ```
   The PNG lands in `decks/<slug>/render/slide-NN.png`; the inventory maps slide ids to numbers.
5. **Review** the PNG and fix what is wrong (next section). At most three rounds.
6. **Show** the slide and **wait**. Do not start another slide until the user reacts.
7. **Revise** from feedback by going back to step 1. On approval set the storyline entry's status
   to `approved` (if there is a storyline) and offer the next open entry.

A slide can also be designed without a storyline, from a single prompt; skip the storyline parts.

To change a slide, change its shapes (`set_text`, `update_shape`, `delete_shape`). Never delete
and rebuild an existing slide: the user's hand edits would be lost. For a variant, duplicate the
slide with `clone_slide` from `"self"` and change the copy.

## The render-check-review loop

Each round: apply → `pptx check` → `pptx render` → view the PNG with your image-viewing tool →
go through [review-checklist.md](references/review-checklist.md) → write the next change set.

- `check` exits with 1 when it finds errors: text overflow, a shape off the slide, leftover
  template text, a font below the minimum. Fix every error before you show the slide. Fix
  warnings too, or say why one stays.
- `check` measures text with the real font; the render comes from LibreOffice, whose line breaks
  can differ slightly from PowerPoint. Leave slack: a box that is filled to the last line in the
  render may overflow in PowerPoint.
- **At most three rounds** of self-correction. Then show the slide anyway and list what remains.

## Show and wait

Report in this order, then stop:

1. the PNG path (on macOS you may `open` it),
2. one line on what the slide argues and how it shows it,
3. open findings: warnings you kept and why, anything unresolved after three rounds, a fallback font.

## Feedback by screenshot

The user sends a screenshot of a slide (from your render or from PowerPoint, marked up or not)
and says what bothers them.

1. **Find the slide:** the slide number in the footer, or the title text matched against
   `pptx inventory`. Ask when it is ambiguous.
2. **Map each complaint to shapes** in the inventory, by name, position (boxes are in cm from the
   top-left corner) and text.
3. **Change those shapes** with a change set, then check, render and review as usual. Compare the
   new render with the screenshot side by side before you show it.

If the screenshot shows a slide that is not in the working copy, ask for the deck it comes from and
import the slide with `clone_slide` (`"from": {"deck": "<their-deck.pptx>", "slide": N}`), then
change it as asked. Without that deck, build a new native slide modeled on it: structure from the
screenshot, style from the brief. Never insert the screenshot as a picture.

## Getting slides back

- The working copy is a normal `.pptx`. The user opens it in PowerPoint and copies approved slides
  into their own deck; you never write to the user's original.
- If the user edits the working copy in PowerPoint, they save before asking you to continue and
  reload after you write. When `apply` notes that the file changed outside the agent, those are
  hand edits: read the new inventory and keep them.
- **Lock file:** before you apply, look for `~$<slug>.pptx` in the deck folder (`ls decks/<slug>`).
  It means the deck is open in PowerPoint. Say so before you write: PowerPoint will not show your
  change until the user reloads, and saving from PowerPoint afterwards would overwrite it.
- **Slides edited in the user's own deck** come back with `replace_slide`, which keeps position,
  slide id, storyline and change log. Match each edited slide to its older version by title
  (`pptx inventory <their-deck.pptx>` next to the working copy's inventory) and ask when unsure;
  then replace it, `"from": {"deck": "<their-deck.pptx>", "slide": N}`. Their text is real
  content, never example text. To start over from their deck instead:
  `pptx init decks/<new-slug> --from <their-deck.pptx> --keep-slides`, then copy `storyline.md` across.

## The design method

### Content

1. **One message per slide.** The action title states it as a full sentence, or two short ones,
   in at most two lines at title size. If the title cannot be written, the slide is not ready.
2. **Kicker** = a short label for the slide's role or chapter ("Classification"), not a sentence.
3. **The body proves the title.** Whatever does not support it goes to the notes or another slide.
4. **The structure carries the meaning.** Pick the form from the argument: matrix for two
   dimensions, process for a sequence, comparison for alternatives, timeline for time, tiles for
   parallel facts, chart for numbers, table for lookup ([visual-forms.md](references/visual-forms.md)).
5. **Little text, structured. Non-negotiable.** A label plus one short line; a bold lead-in
   followed by plain text; no paragraph longer than three lines; no bullet walls. Detail goes into
   the speaker notes (op `set_notes`).
6. **Concrete over abstract.** Examples ("e.g. implement and test a requirement"), not buzzwords.

### Visual

7. **Use the master.** Layout placeholders for kicker, title, footer and number; free shapes only
   inside the content area. Kicker and title come empty from the layout; fill them with `set_text`.
   Footer, slide number and date come from the layout with their text and fields; leave them.
8. **Align to the grid.** Edges line up, gutters are equal, equal roles get equal sizes. Take boxes
   from `pptx grid`; never estimate them.
9. **Color has meaning.** Theme colors only. One color family for value and emphasis, neutral
   fills for structure, a highlight color at most once or twice per slide, and the same meaning in
   the same color across the deck.
10. **Hierarchy through weight and size**, with few sizes per slide (title plus one or two body
    sizes), none below the brief's minimum.
11. **Native and editable.** No pictures of text, diagrams or tables.
12. **Consistent across slides.** Title position, kicker use and footer stay the same.

### Non-negotiable

- **No keyword chains.** Never join items with " · ", " | " or " • " ("Speed · Quality · Scale").
  Give each item its own shape or line with a bold lead-in and a plain explanation, or cut it.
- **Little text.** Labels plus one short line, at most three lines per paragraph, details to the
  notes. When text does not fit, cut it or split the slide; never shrink the font to make it fit.

### Anti-patterns

- keyword chains joined by " · ", " | " or " • "
- too much text: long paragraphs, bullet walls, a font shrunk to make it fit
- topic titles ("Overview", "Background") instead of statements
- decoration without meaning: icons, gradients, shadows, colors the master does not use
- everything in identical cards with no hierarchy
- a "key takeaway" box that repeats the title
- elements almost aligned; uneven gaps
- free text boxes where the layout has a placeholder

## Building with change sets

| Mistake | Instead |
|---|---|
| Added shape with text but no `font`: it gets the master's default text (often 18 pt, centered) | Give every text shape `font` (size, color), `align`, `anchor` and `insets` |
| `textbox` for a block of text: it grows with its text, so `check` cannot report overflow | `rect` with `"fill": "none", "line": "none"` |
| Fill or line left out, expecting a default look | Added shapes have no fill and no line until you set them |
| Boxes estimated by eye | `pptx grid decks/<slug> --cols C --rows R`, then copy the boxes |
| Title or kicker as a free text box | `set_text` with `{"role": "title"}` or `{"role": "kicker"}` |
| Generic names like `Rectangle 5` | Names that say what the shape is: `matrix-cell-r2-c3`, `axis-autonomy` |
| python-pptx code or XML edits | Change sets only; `apply` keeps a snapshot and a log |

## Red flags: stop and correct

- You are about to show a slide whose PNG you have not viewed.
- A cell is getting a third line, or a list of items joined by separators.
- You are designing the next slide before the user reacted to this one.
- The title is a topic ("Market overview"), not a statement.
- You are about to delete and rebuild a slide to change it.

## References

- [changeset-format.md](references/changeset-format.md): every op with examples, colors, units, slide and shape refs.
- [visual-forms.md](references/visual-forms.md): which form for which argument, with `pptx grid` recipes and change-set skeletons.
- [review-checklist.md](references/review-checklist.md): the visual review for every round.
