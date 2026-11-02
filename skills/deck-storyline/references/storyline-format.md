# storyline.md format

`decks/<slug>/storyline.md` is the one file both skills share. `deck-storyline` writes and
restructures it; `slide-design` works through it one entry at a time and updates each entry's
status and id, and its kicker and title when their wording changes during design. Field labels and status values stay in English whatever the deck's language, so
both skills find them; the values (titles, kickers, descriptions) are in the deck's language.

## Header

The file starts with `# Storyline: <deck title>`, followed by these fields as a bullet list
`- **Field:** value`, in this order:

| Field | Value |
|---|---|
| Purpose | What the audience should decide, know or do afterwards, in one sentence |
| Core message | The one sentence the whole deck argues; the summary slide states it |
| Audience | Who they are and what they already know |
| Format | `live talk`, `pre-read`, or both ("live talk, sent as a pre-read afterwards") |
| Duration | Minutes and the slide count that fits them ("20 minutes, 8 slides") |
| Language | The language of all text on the slides |
| Style sources | Paths for `pptx init`: master, example deck, style file. For a deck being continued, its path plus "keep slides". `not chosen yet` when the user has not named one |
| Input | Files and folders the storyline is built from: research run, report, notes; "conversation" for free text |

## Entries

After the header comes `## Slides` with one entry per slide, in the intended slide order:

```text
### Slide <N>
- **Kicker:** <short label>
- **Title:** <action title>
- **Shows:** <the visual idea>
- **Sources:** <where the facts come from>
- **Status:** <open, in design, approved or existing>
- **Id:** <slide id, or —>
- **Notes:** <optional>
```

| Field | Value |
|---|---|
| Kicker | The slide's role or chapter in one to three words ("Summary", "Options", "Pilot results"). Not a sentence. `—` for slides without one, such as the title slide |
| Title | The action title: the slide's message as a full sentence, or two short ones, at most about 120 characters (two lines at title size) |
| Shows | The visual idea in one or two sentences: the form (matrix, process, comparison, timeline, tiles, chart, table) and what it holds. Not the slide's text |
| Sources | Where the facts in title and body come from. From a research run: source id and short name, claim ids in parentheses (`S004 Pilot survey (C011)`); otherwise the file or document ("pilot retro notes"). `—` when the slide states no facts |
| Status | One of the four statuses below |
| Id | The slide's stable id in the working copy (`pptx inventory` prints it as `id 257`); `—` until the slide exists |
| Notes | Optional: detail meant for the speaker notes, what was cut from the slide, open questions starting with "Open question:" |

`<N>` is the entry's position, which is the intended slide number. After an insert, move or cut,
renumber every heading so they run 1, 2, 3, … again. Positions change; ids never do.

## Statuses

| Status | Meaning | Set by |
|---|---|---|
| `open` | To be built (no id) or to be revised (with an id) | `deck-storyline`, for new entries and when the message of a built slide changes |
| `in design` | `slide-design` is building or revising the slide | `slide-design`, when it starts the slide |
| `approved` | The user approved the slide | `slide-design`, on approval |
| `existing` | The slide was in the deck before the storyline, or was added by hand, and has not been redesigned | `deck-storyline`, from `pptx inventory --outline` |

When the user changes the kicker, title or "Shows" of an `approved` or `existing` entry, set it
to `open` and keep its id: `slide-design` then revises that slide instead of building a new one.

## Slide ids

The id is the slide's `sldId` in the working copy. It stays the same when other slides are
inserted, moved or deleted, so an entry always finds its slide. Change sets address it directly
as `"id:<Id>"`. `check` and `render` take slide numbers; the current number for an id is in

```bash
pptx inventory decks/<slug>/<slug>.pptx --outline
```

which prints one line per slide, `<number> (id <id>) | <kicker> | <title>`, with `—` where a slide
has no kicker or title. Each line of an existing deck becomes one entry: number → heading, id →
Id, kicker and title as they are, status `existing`.

## Example

A storyline in the middle of the work: slide 1 came from the deck being continued, slide 2 is
approved, slide 3 is being designed, the rest are still open.

```markdown
# Storyline: Winning back focus time

- **Purpose:** The leadership team approves a two-quarter pilot of meeting-free afternoons in two departments.
- **Core message:** Meeting-free afternoons gave the pilot teams five more focus hours a week without slowing decisions, so we should extend the pilot to two departments.
- **Audience:** Leadership team of a 400-person product organisation; they know the engagement survey, not the pilot results.
- **Format:** live talk with discussion, sent as a pre-read afterwards
- **Duration:** 20 minutes plus 10 minutes of discussion, 8 slides
- **Language:** English
- **Style sources:** `~/Templates/master.potx`; examples from `~/Decks/quarterly-review.pptx`; continues `~/Decks/focus-time-draft.pptx` (keep slides)
- **Input:** `research/runs/20261028-focus-time/` (verified claims and sources); pilot retro notes

## Slides

### Slide 1
- **Kicker:** —
- **Title:** Winning back focus time
- **Shows:** Title slide of the draft deck, unchanged.
- **Sources:** —
- **Status:** existing
- **Id:** 256

### Slide 2
- **Kicker:** Summary
- **Title:** Meeting-free afternoons gave the pilot teams five more focus hours a week, and decisions did not slow down.
- **Shows:** Three tiles: focus hours before and after, decision lead time before and after, and what the teams changed. Each tile has a bold lead-in and one line of evidence.
- **Sources:** pilot retro notes; S004 Pilot survey (C011)
- **Status:** approved
- **Id:** 259

### Slide 3
- **Kicker:** Problem
- **Title:** Engineers now spend more than half of their week in meetings, up from a third three years ago.
- **Shows:** Column chart of meeting hours per week by role over three years, with the engineers' columns in the emphasis color.
- **Sources:** S001 Engagement survey (C002, C003)
- **Status:** in design
- **Id:** 260
- **Notes:** The split by department goes into the speaker notes.

### Slide 4
- **Kicker:** Problem
- **Title:** Fragmented calendars cost more than the meetings themselves, because a 30-minute gap is rarely used for deep work.
- **Shows:** Timeline of a typical engineer's day with meetings and gaps; gaps under one hour are hatched as lost time.
- **Sources:** S002 Calendar study (C005); S006 Interruption research (C006)
- **Status:** open
- **Id:** —

### Slide 5
- **Kicker:** Options
- **Title:** Of three ways to cut meeting load, only meeting-free afternoons protect long focus blocks without new rules for every team.
- **Shows:** Comparison of meeting budgets, shorter default meetings and meeting-free afternoons on focus blocks gained, effort to introduce and risk to decisions; the recommended option highlighted.
- **Sources:** S002 Calendar study (C007); S005 Practitioner survey (C009)
- **Status:** open
- **Id:** —

### Slide 6
- **Kicker:** Pilot results
- **Title:** In the eight-week pilot, six teams kept 85 % of their afternoons free and reported fewer late decisions.
- **Shows:** Bar chart of free afternoons per team, with one line below on how decision lead time changed.
- **Sources:** pilot retro notes
- **Status:** open
- **Id:** —

### Slide 7
- **Kicker:** Risks
- **Title:** Two risks need a rule before the rollout: customer escalations and teams across time zones.
- **Shows:** Two cards, each with the risk, an example from the pilot and the rule that handles it.
- **Sources:** pilot retro notes; S007 Remote collaboration study (C014)
- **Status:** open
- **Id:** —
- **Notes:** Open question: does the support team count as a team across time zones?

### Slide 8
- **Kicker:** Decision
- **Title:** We ask for approval to run the pilot in two departments for two quarters, with a review in week 12.
- **Shows:** Three-step process (prepare, run, review) along a time axis, with the decision needed today next to it.
- **Sources:** —
- **Status:** open
- **Id:** —
```
