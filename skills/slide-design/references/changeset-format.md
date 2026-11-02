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
