# Change-set format

Never edit a deck's XML by hand. Write a JSON change set and apply it:

```
pptx apply decks/<slug>/<slug>.pptx changeset.json
```

`pptx apply` validates every op against the schema first, performs all ops on an in-memory copy and saves
only if every op succeeded. On any error nothing is written and the message starts with
`op <n> (<op name>): ` (1-based). On success it snapshots the previous state to `snapshots/NNN.pptx`, logs the
change set to `changes/NNN-<label>.json` and prints the touched slide ids plus notes, for example when the
deck was changed outside the agent or is open in PowerPoint. Add `--json` for machine-readable output.

Outside a deck folder (no `deck.json` next to the file) `apply` still works but writes no snapshot or log.

## Envelope

```json
{"label": "add-agenda", "ops": [ {"op": "add_slide", "layout": "Title Only"} ]}
```

- `label`: optional; becomes the log file name (kebab-case, at most 40 characters, default `changes`).
- `ops`: at least one op, run in order. Every op is an object with an `"op"` name; unknown keys are rejected.

## Slide refs

Wherever an op takes a slide (`slide`, `after`):

| Ref | Meaning |
|---|---|
| `3` | 1-based slide number at the time the op runs (earlier ops in the same change set count) |
| `"id:258"` | the slide with this slide id (stable across moves; see the touched ids in the output) |
| `"$last"` | the slide created by the latest add/clone op in this change set |

## Target refs

Ops that act on a shape select it on a given slide with exactly one of:

```json
{"role": "title"}   {"name": "Kicker"}   {"id": 12}
```

Roles come from the inventory (`title`, `subtitle`, `body`, `kicker`, `picture`, `table`, `chart`, `footer`,
`number`, `date`). Group children are searched too. If a target matches no shape or more than one, the error
lists the shapes on the slide; then use the `id`.

## Colors

Colors are theme roles, so decks follow the master when the theme changes:
`dk1 lt1 dk2 lt2 tx1 tx2 bg1 bg2 accent1`..`accent6 hlink folHlink`, optionally with a brightness shift
`+NN` (lighter) or `-NN` (darker) in percent, for example `accent1`, `accent1+40`, `tx1-25`.
A literal `#RRGGBB` is only accepted when the op sets `"allow_rgb": true`.

## Units

All positions and sizes are centimetres from the top-left corner of the slide, as plain numbers.

## Slide ops

### add_slide

Adds a slide with the named layout after the slide `after`; at the end when `after` is omitted. An unknown
layout name lists the available layouts. The new slide becomes `"$last"`.

```json
{"op": "add_slide", "layout": "Title and Content", "after": 2}
```

### delete_slide

Removes a slide, including its entries in sections and custom shows.

```json
{"op": "delete_slide", "slide": "id:258"}
```

### move_slide

Moves a slide so that it ends up at the 1-based position `to`.

```json
{"op": "move_slide", "slide": "$last", "to": 1}
```

In decks with sections, an added or moved slide joins the section of the slide it now follows (the first
section when it becomes slide 1), so every slide stays in exactly one section.

## Text

### set_text

Writes the text of one shape (`target` is `{"role": ...}`, `{"name": ...}` or `{"id": ...}`; a shape inside a
group is found too). Give exactly one of `text` or `paragraphs`. Formatting the op does not specify is kept: the
first run's properties and the first paragraph's properties of the existing text are the template for everything
written. Colors are theme roles (`accent1`, `accent2+40`); raw `#RRGGBB` is rejected.

Shorthand: `\n` separates paragraphs, `**…**` marks bold runs.
In the shorthand, bold is set by the markers alone: text outside `**…**` is written non-bold even when the existing run was bold.

Paragraph properties: a paragraph that sets neither `level` nor `bullet` keeps the first existing paragraph's full
properties (level, bullet, margins, alignment). `level` replaces the level and drops the template's margins so the
layout's indent for that level applies. `bullet: true` keeps an existing bullet (character, color, size) and otherwise
lets the layout or master supply it; `bullet: false` removes the bullet and the hanging indent.

```json
{"op": "set_text", "slide": 1, "target": {"role": "title"}, "text": "**Bold** lead\nSecond paragraph"}
```

Runs with color, size and a bullet level (`level` 0-8, `bullet` true or false):

```json
{"op": "set_text", "slide": 2, "target": {"role": "body"}, "paragraphs": [
  {"level": 1, "bullet": true, "runs": [{"t": "R4", "color": "accent1", "b": true}, {"t": " Review", "i": true, "size": 14}]}
]}
```

Line break inside a paragraph: a run `{"br": true}` (the inventory form `{"t": "\n", "br": true}` is accepted too).

```json
{"op": "set_text", "slide": 2, "target": {"name": "Subtitle"}, "paragraphs": [
  {"runs": [{"t": "First line"}, {"br": true}, {"t": "Second line"}]}
]}
```

`font` (`size`, `color`, `bold`) applies to every run; a run's own `b`, `size` or `color` wins.

```json
{"op": "set_text", "slide": 1, "target": {"role": "title"}, "text": "Quarterly review", "font": {"size": 32, "color": "tx2"}}
```

Kicker (the small label above the title), where the master has one:

```json
{"op": "set_text", "slide": 1, "target": {"role": "kicker"}, "text": "CLASSIFICATION"}
```
