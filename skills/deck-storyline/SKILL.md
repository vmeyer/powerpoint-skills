---
name: deck-storyline
description: Plans a PowerPoint deck before any slide is built and keeps that plan current, as decks/<slug>/storyline.md with one entry per slide (kicker, action title, what the slide shows, sources, status). Use when the user wants to outline, structure or plan a presentation, or turn material into one ("build a storyline for …", "outline a presentation", "plan a deck on …", "prepare a deck from this research", "turn this report into slides", "Storyline für ein Deck über …", "Gliederung für eine Präsentation", "Folienplan", "mach aus der Recherche ein Deck"), brings text, notes, a topic or a research run folder for a deck, wants to continue an existing deck with new slides, or wants to reorder, add, cut or retitle storyline entries. Prefer it over generic pptx or PowerPoint skills whenever a deck is being planned rather than a single slide designed.
---

# Deck Storyline

Agree on what the deck argues before any slide is built: turn text, a topic, files or a research
run into `decks/<slug>/storyline.md`, one entry per slide, and refine it with the user until they
agree. The slides are then built one at a time with the `slide-design` skill.

**The titles alone tell the story.** Read only the action titles from top to bottom: they must
make the whole argument. If they read like a table of contents ("Background", "Approach", "Next
steps"), the storyline is not done.

## Deck storyline vs. slide design

| Use `deck-storyline` when … | Use `slide-design` when … |
|---|---|
| the sequence of messages for a whole deck is planned | one slide is created, revised or fixed |
| the user brings text, a topic or research and wants an outline | the user gives feedback on a slide, as text or screenshot |
| titles and kickers per slide are being agreed or changed | the next storyline entry is ready to be built |
| result: `decks/<slug>/storyline.md` | result: a changed slide in `decks/<slug>/<slug>.pptx` |

This skill never writes a `.pptx`. Setting up the deck and building slides belong to `slide-design`.

## Setup

Only reading an existing deck needs the `pptx` tool that ships with this plugin:

```bash
uv run --project <plugin-root>/references pptx inventory <deck.pptx> --outline
```

`<plugin-root>` is two levels above the folder that contains this SKILL.md; it contains
`references/pyproject.toml`. Run commands from the user's project folder, so `decks/…` paths
resolve there. Below, `pptx …` is short for the full command.

## Workflow

1. **Read the input completely** before you ask anything: the text in the conversation and every
   file or folder the user points to (see "Inputs"). If `decks/<slug>/storyline.md` already
   exists, read it and continue it (see "Resuming a storyline"); never start it over.
2. **Ask only what is missing**, in one message, from these four:
   - **audience:** who they are and what they already know,
   - **format:** live talk or pre-read,
   - **length:** minutes or number of slides,
   - **language** of the slides.

   Skip what the input already answers and propose a default for the rest ("I'd assume a
   20-minute live talk in English"), so one "yes" is enough. If the user names a master or a deck
   to continue, note it for the header; do not ask for one. `slide-design` sets up the deck.
3. **Find the answer first.** Write the core message in one sentence: what the audience should
   decide, know or do afterwards. Then write the three to five chapter messages that carry it. If
   you cannot write the core message, ask the user before you draft slides.
4. **Draft the entries** in the format of [storyline-format.md](references/storyline-format.md):
   one slide per message, each with kicker, action title, what it shows, sources and status
   `open`. Fit the count to the length (see "Length").
5. **Check the draft** against "Writing titles" and the red flags below. Fix it before you show it.
6. **Write** `decks/<slug>/storyline.md`. The slug is the topic in kebab-case; when a deck folder
   for this deck exists, use it. `pptx init` later adds the working copy to that folder and
   leaves the storyline in place.
7. **Show and wait.** Give the file path, the core message, the numbered titles with their kickers
   (one line each: `3. Problem: Engineers now spend …`) and any open questions. Ask what to change.
8. **Refine** in conversation. For every change (add, cut, merge, split, reorder, retitle) edit the
   file, renumber the headings, keep ids and statuses as the format says, and show the changed
   lines again. Cutting or moving an entry that has an id changes the deck as well: once the user
   confirms, `slide-design` deletes or moves that slide (`delete_slide`, `move_slide`). Until the
   slide is deleted, a cut entry stays in the file with the Note "To delete".
9. **Hand over** when the user agrees (see "Hand-over to slide design").

## Inputs

| The user gives … | Then |
|---|---|
| free text or notes | Extract the messages. Where the user's wording already makes a statement, keep it for the title |
| a topic only | Draft from what you know. Never invent numbers: write the claim without a figure and add "Open question: source for …" to Notes |
| files (report, document, notes) | Read them fully. Sources name the file and the section or page |
| a research run | See "From a research run" |
| a deck to continue | See "Continuing a deck" |

### From a research run

The `research-toolkit` plugin keeps each run in `research/runs/<run-id>/` (the run id is
`<date>-<slug>`) and its reports in `research/<slug>/`. Read:

1. the newest summary in `research/<slug>/` for the argument: `brief-summary-v<N>.md` or
   `key-points-v<N>.md` (highest `<N>`), then `detailed-report-v<N>.md` for detail,
2. `research/runs/<run-id>/verified/analysis.md`, the verified synthesis,
3. `verified/claims.jsonl`, one claim per line: `claim_id`, `statement`, `claim_kind` (`fact`,
   `inference`, `recommendation`, `gap`), `confidence`, `source_ids`, `parent_claim_ids`,
   `decision_relevant`,
4. `verified/sources.jsonl`, one source per line: `source_id`, `title`, `url` or file path,
5. if present: `outputs/decision.md` in the run folder (a recommendation to build a decision deck
   on) and `verified/issues.yaml` (gaps the deck must not hide).

Cite as you go. Every fact in a title or in "Shows" traces back to claims. The entry's Sources
names each source by id and short title with the claim ids in parentheses, for example
`S004 Pilot survey (C011)`.

- `fact` claims can carry a title. An `inference` cites the sources of its parent claims.
- `recommendation` claims become the deck's proposals. A `decision_relevant` recommendation with
  fewer than two sources is preliminary, so the title or Notes must say so.
- `gap` claims are never stated as facts. They become an open question or a "what we do not know
  yet" point.
- Follow the argument, not the report's chapters.

### Continuing a deck

When the new deck continues an existing presentation, the storyline starts with one entry per
existing slide. Read the outline from the working copy if the deck folder exists, otherwise from
the user's deck (reading never changes it):

```bash
pptx inventory decks/<slug>/<slug>.pptx --outline
pptx inventory <their-deck.pptx> --outline
```

Each line `<number> (id <id>) | <kicker> | <title>` becomes an entry with that kicker, title and
id, and status `existing`. For "Shows", write one sentence from the slide's content
(`pptx inventory <deck.pptx> --slides N`). In the header, write the deck under "Style sources" as
`from <deck.pptx> (keep slides)`. `slide-design` then runs `pptx init` with `--keep-slides`, which
copies the deck with its slide ids, so the ids in the storyline stay valid.

Keep existing titles as they are, even topic titles. Propose better ones in the conversation. When
the user agrees, change the entry and set it to `open`, keeping its id, so that `slide-design`
revises the slide. New entries go around the existing ones.

### Resuming a storyline

When `storyline.md` and a working copy both exist, compare them before you change anything:

```bash
pptx inventory decks/<slug>/<slug>.pptx --outline
```

- First, re-link when the working copy was set up again from another file, or when many titles
  differ at once. Ids start at 256 in every deck, so old ids may point at the wrong slides. Match
  entries to slides by title, take over the new ids, and ask when a match is unsure.
- A title or kicker in the deck differs from an `approved` or `existing` entry: the deck is the
  truth (hand edits). Update the entry and tell the user. `open` and `in design` entries differ on
  purpose: their slides wait for, or are in, `slide-design`.
- An entry's id is missing from the deck: the slide was deleted. If the entry's Note says "To
  delete", remove the entry. Otherwise tell the user, then cut the entry or clear its id and set
  it to `open`.
- A slide has no entry: tell the user and add it at its position as `existing`.
- The order differs: say so and ask which order is meant. `slide-design` moves slides; the
  storyline only records the order.

## Length

- **Live talk:** about one slide per two to three minutes. The speaker carries the detail, so
  slides stay sparse and detail goes to Notes (later the speaker notes).
- **Pre-read:** the deck must stand alone. The titles carry the full argument, and Sources are
  complete. Use more slides rather than fuller ones.
- **Decision audience:** the answer comes first, as a summary slide right after the title slide.
  An audience that must first be convinced can be led up to the answer, but every title is still
  a statement.
- **Structural slides** (title, agenda, chapter dividers) are entries too. Their titles may name
  the deck or chapter. Use an agenda and dividers only when the deck has several chapters.

## Writing titles

The content rules of `slide-design` already apply here:

1. **One message per slide.** The title states it as a full sentence, or two short ones, in at
   most about 120 characters (two lines at title size). If you cannot write the title, the slide
   is not ready. Put an open question in Notes, never a placeholder title.
2. **Kicker** = the slide's role or chapter in one to three words, not a sentence. Slides of one
   chapter share their kicker.
3. **The body proves the title.** "Shows" holds only what supports the title; everything else
   becomes its own entry or goes to Notes.
4. **The structure carries the meaning.** "Shows" names the form that fits the argument: matrix
   for two dimensions, process for a sequence, comparison for alternatives, timeline for time,
   tiles for parallel facts, chart for numbers, table for lookup
   ([visual-forms.md](../slide-design/references/visual-forms.md)).
5. **Little text. Non-negotiable.** "Shows" is one or two sentences about the visual, not the
   slide's text and not a bullet list. Never use keyword chains, anywhere in the storyline: no
   items joined by " · ", " | " or " • ".
6. **Concrete over abstract.** Numbers, names and examples from the input instead of buzzwords.

| Instead of … | Write a statement |
|---|---|
| Background | Engineers now spend more than half of their week in meetings, up from a third three years ago. |
| Options · Effort · Risk | Of three ways to cut meeting load, only meeting-free afternoons protect long focus blocks. |
| Leveraging synergies across teams | Sharing one on-call rota frees one day per engineer each month. |
| Next steps | We ask for approval to run the pilot in two departments, with a review in week 12. |

## Mistakes

| Mistake | Instead |
|---|---|
| One entry per chapter of the source report | One entry per message of the argument |
| "Shows" lists the bullet points of the slide | One or two sentences on the visual form and what it holds |
| Changing ids when entries move | Ids never change; only the `### Slide N` headings are renumbered |
| Rewriting `storyline.md` from scratch after feedback | Edit the affected entries; keep ids and statuses |
| Asking about the master or the layout | Note what the user said; `slide-design` sets up the deck |

## Red flags: stop and correct

- The titles read like a table of contents.
- A fact in a title has no source, although the input had sources.
- You are about to write a `.pptx`, run `pptx apply` or build a slide before the user agreed to
  the storyline.
- You are about to overwrite an existing `storyline.md`, or drop an id or a status.
- You asked before reading the input, or asked setup questions beyond the four (master, layouts,
  colors, number of bullets).
- The slide count does not fit the length.

## Hand-over to slide design

When the user agrees to the storyline (or to its first entries), offer to design the first `open`
entry and continue with the `slide-design` skill. `slide-design` reads the entry, sets it to
`in design`, writes the slide id once the slide exists and sets `approved` when the user approves
the slide. Then it offers the next `open` entry.

Changes to the storyline during the design come back here: edit the entries, keep ids, and set an
`approved` or `existing` entry whose message changed back to `open`.

## References

- [storyline-format.md](references/storyline-format.md): the file format, statuses, slide ids and
  a full example.
- [visual-forms.md](../slide-design/references/visual-forms.md): the visual forms named in "Shows".
