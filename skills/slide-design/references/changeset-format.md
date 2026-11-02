# Change sets

A change set is a JSON file with a list of ops: the quick, precise way to add slides, write texts and
add or change shapes, tables, charts, pictures and notes. Never edit a deck's XML by hand.

```bash
pptx build decks/<slug>/<slug>.pptx decks/<slug>/step.json   # apply, check, render
pptx apply decks/<slug>/<slug>.pptx decks/<slug>/step.json   # apply only; then `pptx check decks/<slug>`
```

## Examples first

A new slide: title, a statement block, three numbered steps, notes.

```json
{"label": "three-steps", "ops": [
  {"op": "add_slide", "layout": "Title Only"},
  {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "Three steps take a team from pilot to routine"},
  {"op": "add_shape", "slide": "$last", "shape": "rect", "box": [1.7, 4.5, 9, 12.5], "fill": "accent1",
   "text": "**Start small:** one team, one process, six weeks", "font": {"size": 24, "color": "bg1"}, "anchor": "bottom"},
  {"op": "add_shape", "slide": "$last", "shape": "ellipse", "box": [12.5, 5, 1.6, 1.6], "fill": "accent2",
   "text": "**1**", "font": {"color": "bg1"}, "align": "center", "anchor": "middle", "insets": [0, 0, 0, 0]},
  {"op": "add_shape", "slide": "$last", "shape": "rect", "box": [14.6, 4.9, 17.5, 2], "text": "**Measure today:** lead time, defects, waiting"},
  {"op": "set_notes", "slide": "$last", "text": "The steps repeat for every further team."}
]}
```

Revising: change a text, recolor and move a shape, rename it.

```json
{"label": "revise", "ops": [
  {"op": "set_text", "slide": 3, "target": {"role": "title"}, "text": "Two steps are enough for the first team"},
  {"op": "update_shape", "slide": 3, "target": {"name": "rect-1"}, "fill": "#0B5CAD", "box": [1.7, 4.5, 12, 12.5], "name": "statement"},
  {"op": "delete_shape", "slide": 3, "target": {"name": "ellipse-1"}}
]}
```

## Rules and defaults

- **Envelope:** `{"label": "…", "ops": [ … ]}`. `label` names the log file (`changes/NNN-<label>.json`);
  ops run in order. Unknown ops and keys are rejected, and any error changes nothing; the message
  starts with `op <n> (<op>): `.
- **Names** are optional. Without `name`, an added shape is called `<kind>-<n>` (`rect-1`, `textbox-2`,
  `line-1`, `table-1`, `chart-1`, `picture-1`, `group-1`); the output lists the names given. A name you
  give must be new on the slide.
- **Text in an added shape** gets, unless the op says otherwise: the master's smallest body size (14 pt
  without a master), color `tx1`, left, top, insets 0.2 cm. Added shapes have **no fill and no line**
  until you set them. A `textbox` grows with its text; a `rect` keeps its size, so `check` reports overflow.
- **Colors:** theme roles `dk1 lt1 dk2 lt2 tx1 tx2 bg1 bg2 accent1`…`accent6 hlink folHlink`, with an
  optional brightness `+NN` (lighter) or `-NN` (darker) in percent (`accent1+40`, `tx1-25`); any `#RRGGBB`;
  `none` for no fill, line or explicit run color. Theme roles follow `pptx theme` and the master.
- **Units:** cm from the slide's top-left corner; a box is `[x, y, w, h]`, a point `[x, y]`. A 16:9 slide
  is 33.87 × 19.05 cm. With a master, boxes and points snap to its guides and content-area edges within
  `snap_cm` (0.05 cm).
- **Slide refs** (`slide`, `after`): a 1-based number, `"id:<n>"` (stable, never reused), or `"$last"`,
  the slide created by the latest `add_slide`, `clone_slide` or `replace_slide` of this change set.
- **Target refs** select one shape on the slide, also inside groups: `{"role": "title"}` (roles `title`,
  `subtitle`, `body`, `kicker`, `picture`, `table`, `chart`, `footer`, `number`, `date`, `other`),
  `{"name": "…"}` or `{"id": 12}`. When several match, the error lists them; use the `id`.
- `pptx inventory decks/<slug>/<slug>.pptx --slides N` shows slides (`## Slide <n> (id <id>)`), shapes,
  names, boxes and texts; `--outline` gives one line per slide.

## Ops

| Op | Does |
|---|---|
| `add_slide` | adds a slide with a layout (`layout`, optional `after`) |
| `delete_slide`, `move_slide` | removes a slide; moves one to the 1-based position `to` |
| `clone_slide` | copies a slide (`from: {deck, slide}`, optional `after`) |
| `replace_slide` | puts a copied slide in place of `slide`, keeping position and id |
| `set_text` | writes the text of a placeholder or shape |
| `add_shape`, `update_shape` | adds a shape; moves, resizes, restyles, renames or re-texts one |
| `add_line` | a straight line with optional arrowheads |
| `delete_shape`, `z_order`, `group`, `ungroup` | removes, restacks, groups or ungroups shapes |
| `add_table`, `set_table` | a native table; new texts for an existing one |
| `add_chart` | a native, editable chart colored from the theme |
| `add_picture` | a PNG or JPG cropped to fill a box or a picture placeholder |
| `set_notes` | the speaker notes |

### Slides

`add_slide` takes a layout name from `style/brief.md` (an unknown name lists the available ones). Title and
kicker come empty; footer, date and slide number come from the layout. In decks with sections, a new or moved
slide joins the section of the slide before it.

```json
{"label": "new-opener", "ops": [
  {"op": "add_slide", "layout": "Title Only"},
  {"op": "move_slide", "slide": "$last", "to": 1},
  {"op": "delete_slide", "slide": "id:258"}
]}
```

`clone_slide` and `replace_slide` copy a slide with its shapes, pictures, charts, tables and notes.
`from.deck` is `"examples"` (the deck given to `init --examples`), `"self"` (a duplicate) or a path to any
`.pptx`/`.potx`. Text cloned from the examples deck counts as leftover until you replace it (`check`:
`leftover-text`). Embedded objects, video, audio and SmartArt cannot be copied; the error says so.
`replace_slide` brings back a slide the user edited in their own deck; links to the old slide follow.

```json
{"label": "variant", "ops": [
  {"op": "clone_slide", "from": {"deck": "self", "slide": 3}, "after": 3},
  {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "The second variant halves the cost"},
  {"op": "replace_slide", "slide": "id:260", "from": {"deck": "board-deck.pptx", "slide": 4}}
]}
```

### Text

`set_text` takes `text` or `paragraphs` and keeps formatting it does not set. In `text`, `\n` starts a
paragraph and `**…**` marks bold runs (the other runs are then not bold). `paragraphs` are
`[{"level"?, "bullet"?, "runs": [{"t", "b"?, "i"?, "size"?, "color"?} | {"br": true}]}]`; `{"br": true}` is
a line break inside a paragraph. `font` (`size`, `color`, `bold`) applies to every run; a run's own value
wins. `bullet: false` removes bullet and indent, `true` keeps or restores the layout's bullet.

```json
{"op": "set_text", "slide": 1, "target": {"role": "body"}, "paragraphs": [
  {"bullet": true, "runs": [{"t": "Faster: ", "b": true}, {"t": "lead time down by half"}]},
  {"runs": [{"t": "Same quality,"}, {"br": true}, {"t": "no extra defects", "color": "accent2"}]}
]}
```

### Shapes

`add_shape` takes `shape`: `rect`, `roundRect`, `ellipse`, `textbox`, `homePlate` (pentagon arrow),
`chevron`, `rightArrow`, `leftArrow`, `upArrow`, `downArrow`; `box`; and optionally:

| Field | Value |
|---|---|
| `fill` | a color or `none` |
| `pattern` | `{"preset", "fg", "bg"}` with OOXML names such as `wdUpDiag`, `ltUpDiag`, `pct20`; not with `fill` |
| `line` | `"none"` or `{"color", "width_pt", "dash"}`; `dash`: `solid dash sysDash dot sysDot lgDash dashDot` |
| `text` or `paragraphs`, `font` | as in `set_text` |
| `align`, `anchor`, `insets` | `left center right`; `top middle bottom`; `[l, t, r, b]` in cm |
| `radius` | corner rounding of a `roundRect`, 0–0.5 of the shorter side |
| `rotation` | degrees clockwise about the box centre; `box` is the frame before the turn |

`update_shape` takes the same fields (not `shape`) plus `name`; `font` alone restyles the existing runs and
`"rotation": 0` turns a shape back. A line accepts only `name` and `line`. `add_line` needs `from`, `to`,
`color`, `width_pt`; `dash`, and `head`/`tail` (`none`, `triangle`, `arrow`) are optional. `z_order` takes
`to`: `front`, `back`, `forward`, `backward`. `group` takes two or more top-level `targets`. Deleting or
ungrouping removes the animations of what is gone; `apply` notes it.

```json
{"label": "path", "ops": [
  {"op": "add_shape", "slide": 1, "shape": "homePlate", "box": [1.7, 4.5, 10, 1.4], "fill": "accent1",
   "text": "**Assess**", "font": {"color": "bg1"}, "anchor": "middle"},
  {"op": "add_shape", "slide": 1, "shape": "chevron", "box": [11.4, 4.5, 10, 1.4], "fill": "accent1+40",
   "text": "**Pilot**", "anchor": "middle", "align": "center"},
  {"op": "add_line", "slide": 1, "from": [1.7, 8], "to": [32.2, 8], "color": "tx1+50", "width_pt": 1.5, "tail": "triangle"},
  {"op": "add_shape", "slide": 1, "shape": "rect", "box": [1.7, 9, 9, 4], "pattern": {"preset": "wdUpDiag", "fg": "accent6+60", "bg": "bg1"}},
  {"op": "z_order", "slide": 1, "target": {"name": "line-1"}, "to": "front"}
]}
```

### Tables, charts, pictures, notes

`add_table`: `box`, `rows` (cell texts with the `text` shorthand), optional `header` (default true),
`col_widths` (cm), `style` (a name from the brief's table styles or a GUID; default: the master's default),
`font`, `fills` (`{"row,col": color}`, 0-based). `set_table` writes new rows into an existing table,
keeping its style and the last row's formatting; the number of columns must stay.

`add_chart`: `box`, `type` (`column_clustered`, `column_stacked`, `bar_clustered`, `bar_stacked`, `line`,
`doughnut`), `categories`, `series` (`[{"name", "values"}]`, one value per category; a doughnut takes one
series), optional `legend` (`bottom`, `right`, `none`) and `labels`. Series take `accent1` to `accent6`.

`add_picture`: `path` (PNG or JPG, relative to the current folder), `alt` (one sentence), and `box` (cropped
to fill) or `target` (a picture placeholder). `set_notes`: `text`, `\n` starts a paragraph.

```json
{"label": "evidence", "ops": [
  {"op": "add_chart", "slide": 2, "box": [1.7, 4.5, 20, 12.5], "type": "column_clustered",
   "categories": ["Q1", "Q2", "Q3", "Q4"], "series": [{"name": "Lead time (days)", "values": [42, 35, 28, 21]}], "labels": true},
  {"op": "add_table", "slide": 3, "box": [1.7, 4.5, 30.5, 5], "rows": [["Option", "Cost"], ["Buy", "Low"], ["Build", "**High**"]]},
  {"op": "add_picture", "slide": 4, "path": "assets/team.jpg", "box": [17, 0, 16.87, 19.05], "alt": "The team at the planning workshop"},
  {"op": "set_notes", "slide": 2, "text": "Lead time fell in every quarter.\nAsk who owns the remaining wait."}
]}
```

## Grid

`pptx grid decks/<slug> --cols 3 --rows 2` prints evenly spaced boxes inside the content area (`--gutter`,
`--row-gutter`, `--area full` or `--area x,y,w,h`; `--json` for lists), handy when several shapes must line up.
