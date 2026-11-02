# Change-set format

Never edit a deck's XML by hand. Write a JSON change set (for example `decks/<slug>/change.json`)
and apply it:

```bash
pptx apply decks/<slug>/<slug>.pptx decks/<slug>/change.json
```

`pptx apply` validates every op against the schema first, performs all ops on an in-memory copy and saves
only if every op succeeded. On any error nothing is written and the message starts with
`op <n> (<op name>): ` (1-based). On success it snapshots the previous state to `snapshots/NNN.pptx`, logs the
change set to `changes/NNN-<label>.json` and prints the touched slide ids plus notes, for example when the
deck was changed outside the agent or is open in PowerPoint. Add `--json` for machine-readable output.

Outside a deck folder (no `deck.json` next to the file) `apply` still works but writes no snapshot or log.

## Ops

| Op | Does |
|---|---|
| [`add_slide`](#add_slide) | add a slide with a layout of the master |
| [`delete_slide`](#delete_slide) | remove a slide |
| [`move_slide`](#move_slide) | move a slide to another position |
| [`clone_slide`](#clone_slide) | copy a slide from the examples deck or another deck, or duplicate one |
| [`replace_slide`](#replace_slide) | put a slide from another deck in place of a slide, keeping its position and id |
| [`set_text`](#set_text) | write the text of a placeholder or shape |
| [`add_shape`](#add_shape) | add a rectangle, rounded rectangle, arrow, chevron, ellipse or text box |
| [`add_table`, `set_table`](#add_table-set_table) | add a native table, or fill and resize a cloned one |
| [`add_chart`](#add_chart) | add a native, editable chart colored from the theme |
| [`add_picture`](#add_picture) | add a PNG or JPG cropped to fill a box or a picture placeholder |
| [`set_notes`](#set_notes) | write the speaker notes of a slide |
| [`add_line`](#add_line) | add a straight line, with arrowheads if wanted |
| [`update_shape`](#update_shape) | move, resize, restyle, rename or re-text a shape |
| [`delete_shape`](#delete_shape-z_order-group-ungroup) | remove a shape |
| [`z_order`](#delete_shape-z_order-group-ungroup) | bring a shape to the front or send it back |
| [`group`, `ungroup`](#delete_shape-z_order-group-ungroup) | group shapes or dissolve a group |

A cloned slide already brings its tables, charts and pictures along.

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
| `"id:258"` | the slide with this slide id (stable across moves, and never reused for a new slide in the deck folder, even after a delete; see the touched ids in the output) |
| `"$last"` | the slide created by the latest `add_slide`, `clone_slide` or `replace_slide` in this change set (an error if there is none) |

`pptx inventory <deck.pptx>` lists every slide as `## Slide <number> (id <id>) — <layout>`. A hidden slide is
marked `(hidden)` (`"hidden": true` with `--json`); it keeps its number and renders like any other.

## Target refs

Ops that act on a shape select it on a given slide with exactly one of:

```
{"role": "title"}   {"name": "Kicker"}   {"id": 12}
```

Roles come from the inventory (`title`, `subtitle`, `body`, `kicker`, `picture`, `table`, `chart`, `footer`,
`number`, `date`, `other`). The `kicker` is a short body placeholder above the title, where the layout has one.
Group children are searched too. If a target matches no shape or more than one, the error lists the shapes on
the slide; then use the `id`.

## Colors

Colors are theme roles, so decks follow the master when the theme changes:
`dk1 lt1 dk2 lt2 tx1 tx2 bg1 bg2 accent1`..`accent6 hlink folHlink`, optionally with a brightness shift
`+NN` (lighter) or `-NN` (darker) in percent, for example `accent1`, `accent1+40`, `tx1-25`.
`none` means no fill, no line or (for a run) no explicit color. A literal `#RRGGBB` is only accepted by
`add_shape`, `add_line`, `update_shape`, `add_table` and `set_table` when the op sets `"allow_rgb": true`; `set_text`
never takes one.

## Units

All positions and sizes are centimetres from the top-left corner of the slide, as plain numbers. A box is
`[x, y, w, h]`, a point `[x, y]`.

## Slide ops

### add_slide

Adds a slide with the named layout after the slide `after`; at the end when `after` is omitted. The layout names
are listed in `style/brief.md`; an unknown name lists the available layouts. The new slide becomes `"$last"`.
Kicker and title come empty from the layout; fill them with `set_text`. Footer, slide number and date come
from the layout with their text and fields; leave them.

```json
{"op": "add_slide", "layout": "Title and Content", "after": 2}
```

### delete_slide

Removes a slide, including its entries in sections and custom shows. Links to it from other slides are removed;
`apply` notes which slides had them. A slide that a Slide Zoom on another slide shows cannot be deleted or
replaced: ask the user to remove the zoom in PowerPoint first.

```json
{"op": "delete_slide", "slide": "id:258"}
```

### move_slide

Moves a slide so that it ends up at the 1-based position `to`.

```json
{"op": "move_slide", "slide": "$last", "to": 1}
```

In decks with sections, an added, cloned or moved slide joins the section of the slide it now follows (the
first section when it becomes slide 1), so every slide stays in exactly one section. A replaced slide keeps its
sections.

### clone_slide

Copies a slide into the working copy after the slide `after`; at the end when `after` is omitted. The copy
becomes `"$last"`. `from.slide` is the slide's number (or `"id:<n>"`) in the deck that `from.deck` names:

| `from.deck` | Source |
|---|---|
| `"examples"` | the examples deck given to `pptx init --examples` (numbers as in `style/examples.png` and the building-blocks catalog) |
| `"self"` | the working copy itself: a duplicate; `from.slide` may also be `"$last"` |
| a path | any other `.pptx` or `.potx`, absolute or relative to the current folder |

The copy gets the working copy's layout of the same name and keeps every shape, the background, pictures, charts
with their data, tables, groups, web links and the speaker notes. Links to other slides survive only in a
duplicate, and PowerPoint tags (add-in data) are dropped; `apply` reports both as notes. A missing layout name
stops with the list of available layouts. A slide with an embedded object, video, audio, SmartArt, a chart
whose data is not an embedded `.xlsx` workbook, or (from another deck) a Slide, Section or Summary Zoom stops with
`slide N uses <type>, which cannot be cloned`; build that slide from its layout instead.

```json
{"op": "clone_slide", "from": {"deck": "examples", "slide": 36}, "after": 2}
```

**Example text counts as leftover.** The texts of a slide cloned from the examples deck are recorded, and
`check` reports every shape whose text is still the example's as `leftover-text` (an error). Replace each text
with `set_text`, or delete the shape, before you show the slide. A duplicate keeps that record; slides from other
decks carry real content and are not recorded.

A variant of slide 3, right after it:

```json
{"label": "variant", "ops": [
  {"op": "clone_slide", "from": {"deck": "self", "slide": 3}, "after": 3},
  {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "The second variant halves the lead time"}
]}
```

### replace_slide

Puts a copy of a slide from another deck in place of `slide`: same position, same slide id (so `"id:<n>"` refs
and the storyline stay valid), same sections; links to it from other slides lead to the new version. A slide
that a Slide Zoom shows is refused, as for `delete_slide`. `from` and the copying work as in `clone_slide`.
Use it to bring slides the user edited in their own deck back into the working copy. The replaced version is
gone from the working copy; the snapshot of the apply keeps it. The new slide becomes `"$last"`.

```json
{"op": "replace_slide", "slide": "id:258", "from": {"deck": "board-deck.pptx", "slide": 4}}
```

## Text

### set_text

Writes the text of one shape (`target` is `{"role": ...}`, `{"name": ...}` or `{"id": ...}`; a shape inside a
group is found too). Give exactly one of `text` or `paragraphs`. Formatting the op does not specify is kept: the
first run's properties and the first paragraph's properties of the existing text are the template for everything
written. Colors are theme roles (`accent1`, `accent2+40`); raw `#RRGGBB` is rejected.

Shorthand: `\n` separates paragraphs, `**…**` marks bold runs. With markers, the marked runs are bold and all
other runs explicitly not bold. Without markers no weight is written, and a bold copied from the existing first run
is dropped, so the text takes the weight of the layout or master (a title stays bold where the master makes titles
bold). `font.bold` sets the weight of every run that has no marker.

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

Kicker (the small label above the title), where the master has one. Write it in normal case; a master that
shows kickers in capitals does so by itself.

```json
{"op": "set_text", "slide": 1, "target": {"role": "kicker"}, "text": "Classification"}
```

An action title of two sentences: one sentence per line, with a line break inside one paragraph.

```json
{"op": "set_text", "slide": 1, "target": {"role": "title"}, "paragraphs": [
  {"runs": [{"t": "Three quick wins come first."}, {"br": true}, {"t": "The platform pays off only after the data work"}]}
]}
```

## Shapes

Compose slides from native shapes. Colors are theme roles (`accent1+60`, `dk2-25`, `none`); a raw `#RRGGBB` is
rejected unless the op sets `"allow_rgb": true`. Boxes are `[x, y, w, h]` in cm and snap to the master's grid
(content-area edges and guides, within `rules.snap_cm`) when the deck has a style brief. Every shape gets the
`name` you give it; a name already used on the slide is an error.

### add_shape

`shape` is one of `rect`, `roundRect`, `homePlate` (pentagon arrow), `chevron`, `rightArrow`, `leftArrow`,
`upArrow`, `downArrow`, `ellipse`, `textbox`. Optional fields:

| Field | Value |
|---|---|
| `fill` | a color, or `none` |
| `pattern` | `{"preset", "fg", "bg"}`: OOXML pattern names such as `wdUpDiag`, `ltUpDiag`, `dkDnDiag`, `pct20`; not together with `fill` |
| `line` | `"none"` or `{"color", "width_pt", "dash"}`; `dash` is one of `solid dash sysDash dot sysDot lgDash dashDot lgDashDot lgDashDotDot` |
| `text` or `paragraphs` | as in `set_text`, not both |
| `font` | `{"size", "color", "bold"}` for every run |
| `align` | `left`, `center` or `right` |
| `anchor` | `top`, `middle` or `bottom` |
| `insets` | `[l, t, r, b]` in cm, the space between the shape's edge and its text |
| `radius` | 0-0.5, corner rounding of a `roundRect` |
| `allow_rgb` | `true` to accept `#RRGGBB` |

An added shape shows only what the op sets: no fill, no line, no shadow. Its text gets the master's default
text style (often 18 pt) and, in every shape but `textbox`, is centered horizontally and vertically. So set
`fill` and `line`, and give every shape with text `font` (size and color), `align`, `anchor` and `insets`.

A `textbox` grows with its text, so it never overflows: `check` warns `grows-to-fit` when the text needs more
height than its box and checks overlaps and the slide edge at the grown height. For a block of text with a fixed
size, use a `rect` with `"fill": "none", "line": "none"`; `check` reports its overflow as an error.

Matrix cell with a bold label followed by a plain line:

```json
{"op": "add_shape", "slide": 1, "name": "matrix-cell-r1-c1", "shape": "rect", "box": [1.69, 4.45, 9.96, 6.14],
 "fill": "accent1+60", "line": "none", "text": "**No goal**\nNot worth it yet", "font": {"size": 11, "color": "tx1"},
 "align": "left", "anchor": "top", "insets": [0.3, 0.2, 0.3, 0.2]}
```

Text block without fill or frame:

```json
{"op": "add_shape", "slide": 1, "name": "lead-in", "shape": "rect", "box": [1.69, 4.45, 5.3, 2.6],
 "fill": "none", "line": "none", "text": "**Like a new team member:** the more it takes on, the more guidance it needs",
 "font": {"size": 11, "color": "tx1"}, "align": "left", "anchor": "top", "insets": [0, 0, 0, 0]}
```

Hatched area (pattern preset, foreground and background colors):

```json
{"op": "add_shape", "slide": 1, "name": "hatch", "shape": "rect", "box": [10, 5, 6, 4],
 "pattern": {"preset": "wdUpDiag", "fg": "accent6+60", "bg": "lt1"}, "line": "none"}
```

Axis band with an arrow tip:

```json
{"op": "add_shape", "slide": 1, "name": "axis-value", "shape": "homePlate", "box": [7.29, 4.45, 24.88, 0.9],
 "fill": "accent1", "line": "none", "text": "**Value for customers**", "font": {"size": 12, "color": "lt1"},
 "align": "left", "anchor": "middle", "insets": [0.3, 0, 0.6, 0]}
```

Badge:

```json
{"op": "add_shape", "slide": 1, "name": "badge-1", "shape": "ellipse", "box": [1.5, 3, 1, 1],
 "fill": "accent2", "line": "none", "text": "1", "font": {"size": 12, "color": "lt1", "bold": true}, "align": "center", "anchor": "middle"}
```

### add_line

A straight line from `from` to `to` (points snap like boxes). `color` and `width_pt` are required; `dash` as for
shapes. `head` is the start, `tail` the end; each `none`, `triangle` or `arrow`.

```json
{"op": "add_line", "slide": 1, "name": "path", "from": [10, 8], "to": [14, 8], "color": "dk1", "width_pt": 3, "tail": "triangle"}
```

### update_shape

Takes any `add_shape` field except `shape`, plus `name` to rename. `text` or `paragraphs` replace the text (as in
`set_text`); `font` alone restyles the existing runs. Lines accept only `name` and `line`; delete and re-add a
line to move it. Boxes of shapes inside a group cannot be changed; ungroup first.

```json
{"op": "update_shape", "slide": 1, "target": {"name": "badge-1"}, "box": [2, 3, 1, 1], "fill": "accent3"}
```

### delete_shape, z_order, group, ungroup

`z_order` takes `to`: `front`, `back`, `forward` or `backward`. A group encloses its children exactly; only
top-level shapes that are not placeholders can be grouped. Ungroup keeps every child at its place on the slide.

```
{"op": "delete_shape", "slide": 1, "target": {"name": "hatch"}}
{"op": "z_order", "slide": 1, "target": {"name": "badge-1"}, "to": "front"}
{"op": "group", "slide": 1, "targets": [{"name": "matrix-cell-r1-c1"}, {"name": "badge-1"}], "name": "row-1"}
{"op": "ungroup", "slide": 1, "target": {"name": "row-1"}}
```

## Charts, pictures and notes

### add_chart

Adds a native chart (editable in PowerPoint: right-click, Edit Data). `box` is `[x, y, w, h]` in cm, snapped like shapes.

| Field | Meaning |
|---|---|
| `type` | `column_clustered`, `column_stacked`, `bar_clustered`, `bar_stacked`, `line` or `doughnut` |
| `categories` | category labels; every series needs one value per category |
| `series` | `[{"name", "values"}]`; a doughnut takes exactly one series |
| `legend` | `bottom`, `right` or `none`; default `bottom` for several series or a doughnut, else `none` |
| `labels` | `true` shows the value on every bar, point or segment, default `false` |

Series are colored `accent1` to `accent6` in order, as theme colors (a doughnut colors its categories that way);
more than 6 series or categories is an error. Chart text uses the theme's body font in the text color, at the style
brief's `min_size_pt` (11 pt by default). Value-axis gridlines are thin (0.5 pt) and light (text color, 85 % lighter).
There is no 3-D, gradient, shadow or chart title: the slide title states the message.

```json
{"op": "add_chart", "slide": "$last", "name": "chart-lead-time", "box": [1.69, 4.45, 20.22, 12.57],
 "type": "column_clustered", "categories": ["Q1", "Q2", "Q3", "Q4"],
 "series": [{"name": "Lead time (days)", "values": [42, 35, 28, 25]}], "legend": "none", "labels": true}
```

### add_picture

Adds a PNG or JPG (detected by content; SVG, GIF and the like are rejected, convert them first). `path` is relative
to the current directory or absolute. `alt` is required: one sentence saying what the picture shows. Give either:

- `box`: the picture fills the box exactly. The image keeps its aspect ratio; the overflowing dimension is
  cropped evenly on both sides (crop to fill), never stretched.
- `target`: a picture placeholder of the slide (`{"role": "picture"}` or by name); the picture is inserted
  there, cropped to fill, and takes over the placeholder's position.

`name` becomes the picture's name.

```json
{"op": "add_picture", "slide": "$last", "name": "photo-team", "path": "assets/team.jpg", "box": [22.21, 4.45, 9.96, 12.57], "alt": "The team at the planning workshop"}
```

### set_notes

Replaces the speaker notes of a slide and creates the notes page if there is none. `\n` starts a new paragraph.
Detail that does not belong on the slide goes here.

```json
{"op": "set_notes", "slide": "$last", "text": "Lead time fell by 40 % after the release train.\nAsk who owns the remaining wait."}
```

## Tables

### add_table, set_table

`add_table` adds a native table; `box` is `[x, y, w, h]` in cm (snapped like shapes), `rows` the cell texts, all
rows with the same number of columns. Cell text uses the `text` shorthand of `set_text` (`\n` = paragraph,
`**bold**`): text without `**` writes no weight (a bold copied from the cell's existing run is dropped), so cells
take the weight of the table style, such as a bold header row; text with `**` is bold exactly where marked, and
`font.bold: false` unbolds a whole cell.

| Field | Meaning |
|---|---|
| `header` | first row is the header, default `true` |
| `col_widths` | column widths in cm, one per column; the table is as wide as their sum (default: equal widths in `box`) |
| `style` | table style by name or GUID; names come from "Table styles" in `style/brief.md` (default: the master's default style) |
| `font` | `{size, color, bold}` for every cell |
| `fills` | cell fills `{"row,col": color}`, 0-based, theme colors as elsewhere |

Without `style` the table uses the master's default table style (marked "default" in the brief's table-style
list), so it looks like the master's own tables; python-pptx's built-in "Medium Style 2 - Accent 1" applies only
when the deck defines none. Banded rows are switched off. The style sets borders, fills and header weight; to get
the exact cell formatting of an example table (alignment, sizes), clone its slide and use `set_table`.

`set_table` writes the cell texts of an existing table (`target` as for any shape), usually one cloned from the
examples. It keeps the table style and the cell formatting, adds rows by copying the last row and removes
surplus rows. The number of columns must match; it never adds or removes columns. It refuses to copy or cut
a vertically merged cell (unmerge it in PowerPoint first).

```json
{"op": "add_table", "slide": "$last", "name": "table-options", "box": [1.69, 4.45, 30.48, 6],
 "rows": [["Option", "Cost"], ["Buy", "Low"], ["Build", "**High**"]], "col_widths": [10, 20.48], "font": {"size": 11}}
```

```json
{"op": "set_table", "slide": "$last", "target": {"role": "table"}, "rows": [["Region", "Revenue"], ["North", "4 M"]]}
```

`pptx inventory` lists a table's cell texts as `"table": [[...]]`. Cloned example cells count as example text:
`check` flags them with `leftover-text` until replaced.

## Grid

Do not compute columns or matrix cells by hand. `pptx grid <deck-dir> --cols C --rows R` returns boxes inside the
content area from the style brief; `--gutter` sets the column gutter (default 0.3 cm), `--row-gutter` the row
gutter (default: the column gutter), `--area full` uses the whole slide and `--area x,y,w,h` any other region.
With `--json` the result is `{"area": [x, y, w, h], "rows": [[[x, y, w, h], …], …]}`, one list per row.

For a deck whose content area is x 1.69, y 4.45, w 30.48, h 12.57:

```bash
pptx grid decks/<slug> --cols 3 --rows 2
```

```
1.69,4.45 9.96x6.14 | 11.95,4.45 9.96x6.14 | 22.21,4.45 9.96x6.14
1.69,10.88 9.96x6.14 | 11.95,10.88 9.96x6.14 | 22.21,10.88 9.96x6.14
Boxes are [x, y, w, h] in cm, ready for add_shape.
```

Each entry is a ready `box`:

```json
{"op": "add_shape", "slide": 1, "name": "tile-r1-c1", "shape": "rect", "box": [1.69, 4.45, 9.96, 6.14], "fill": "lt2", "line": "none"}
```

Recipes for matrices, processes, comparisons, timelines and tiles are in `visual-forms.md`.
