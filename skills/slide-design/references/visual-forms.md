# Visual forms

Pick the form from the argument, not from habit. The structure must say what the title claims
before anyone reads the cells.

| The argument is about … | Form |
|---|---|
| two dimensions: where things sit, what combination pays off | [Matrix](#matrix) |
| a sequence: steps, phases, a flow with an outcome | [Process and chevrons](#process-and-chevrons) |
| alternatives: option A vs. B, before vs. after | [Comparison](#comparison) |
| development over time: milestones, a roadmap | [Timeline](#timeline) |
| three to six parallel facts of equal weight | [Tiles](#tiles) |
| numbers: size, trend, share | [Chart](#chart) |
| many values the reader looks up | [Table](#table) |

If none fits, the slide may carry two arguments: split it.

## How to read the recipes

- Every box comes from `pptx grid`. The outputs below are real, from a 16:9 deck whose content
  area is x 1.69, y 4.45, w 30.48, h 12.57 cm (right edge 32.17, bottom 17.02). Your master's
  content area is in `style/brief.md`; run the same commands on your deck and use its numbers.
- `--area x,y,w,h` restricts a grid to part of the content area. Derive it from the content area:
  for a header column of width 5.3 with a gutter of 0.3, the cells start at x + 5.6 and are
  w − 5.6 wide. Shapes that should line up get the same `--cols`, `--gutter`, x and w.
- A shape spanning several cells: x of the first cell, width = right edge of the last − x of the first.
- Colors are roles, not a palette: `accent1` and its tints (`accent1+40`, `accent1+60`) carry value
  and emphasis, `lt2` stands for the neutral structure fill (use the grey your examples use, see
  "Color usage" in the brief), `accent6+80` for the one highlight. Map them to your master.
- Text in a cell: a bold label plus one plain line, both at the brief's minimum size or one step
  above. Every text shape sets `font`, `align`, `anchor` and `insets`, because an added shape
  otherwise shows the master's default text (often 18 pt, centered).
- The skeletons build a new slide on the layout with only kicker and title (`Title Only` in many
  masters; take the name from the brief). Kicker and title are set as in the matrix skeleton. They
  show one op per kind of shape; repeat it for every cell with the boxes from the grid. To revise
  an existing slide, drop the `add_slide` op and use its number or `"id:<n>"` instead of `"$last"`.

## Matrix

**Use when** the argument has two dimensions: a position on two scales (effort and value,
maturity and ambition), or which combination works. **Not** for a list that happens to fit in
a grid: if the rows and columns are not real scales, use tiles.

Structure: an axis band across the columns, column headers, row headers, cells colored by meaning
(tints of one accent for value, the highlight color once for risk or "today", hatching for "not
sensible"). Optional: a two-line lead-in in the top-left corner.

```bash
pptx grid decks/<slug> --cols 3 --area 7.29,5.65,24.88,1.4            # column headers
pptx grid decks/<slug> --cols 3 --rows 3 --area 7.29,7.35,24.88,9.67  # cells
pptx grid decks/<slug> --cols 1 --rows 3 --area 1.69,7.35,5.3,9.67    # row headers, same rows
```

```
7.29,5.65 8.09x1.4 | 15.68,5.65 8.1x1.4 | 24.08,5.65 8.09x1.4

7.29,7.35 8.09x3.02 | 15.68,7.35 8.1x3.02 | 24.08,7.35 8.09x3.02
7.29,10.67 8.09x3.03 | 15.68,10.67 8.1x3.03 | 24.08,10.67 8.09x3.03
7.29,14 8.09x3.02 | 15.68,14 8.1x3.02 | 24.08,14 8.09x3.02

1.69,7.35 5.3x3.02
1.69,10.67 5.3x3.03
1.69,14 5.3x3.02
```

For any content area (x, y, w, h) with a 5.3 cm header column, a 0.9 cm axis band, 1.4 cm column
headers and 0.3 cm gutters:

| Part | Box or `--area` |
|---|---|
| axis band | box `[x+5.6, y, w−5.6, 0.9]` |
| column headers | `--cols 3 --area x+5.6,y+1.2,w−5.6,1.4` |
| cells | `--cols 3 --rows 3 --area x+5.6,y+2.9,w−5.6,h−2.9` |
| row headers | `--cols 1 --rows 3 --area x,y+2.9,5.3,h−2.9` |
| lead-in (optional) | box `[x, y, 5.3, 2.6]` |

```json
{"label": "matrix-effort-value", "ops": [
  {"op": "add_slide", "layout": "Title Only"},
  {"op": "set_text", "slide": "$last", "target": {"role": "kicker"}, "text": "Prioritization"},
  {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "paragraphs": [
    {"runs": [{"t": "Three quick wins come first."}, {"br": true}, {"t": "The platform pays off only after the data work"}]}]},
  {"op": "add_shape", "slide": "$last", "name": "axis-value", "shape": "homePlate", "box": [7.29, 4.45, 24.88, 0.9],
   "fill": "accent1", "line": "none", "text": "**Value for customers**", "font": {"size": 12, "color": "lt1"},
   "align": "left", "anchor": "middle", "insets": [0.3, 0, 0.6, 0]},
  {"op": "add_shape", "slide": "$last", "name": "matrix-col-c1", "shape": "rect", "box": [7.29, 5.65, 8.09, 1.4],
   "fill": "lt2", "line": "none", "text": "**Low**\ne.g. internal reports", "font": {"size": 11, "color": "tx1"},
   "align": "left", "anchor": "top", "insets": [0.3, 0.15, 0.3, 0.15]},
  {"op": "add_shape", "slide": "$last", "name": "matrix-row-r1", "shape": "rect", "box": [1.69, 7.35, 5.3, 3.02],
   "fill": "lt2", "line": "none", "text": "**Low effort**\nOne team, one sprint", "font": {"size": 11, "color": "tx1"},
   "align": "left", "anchor": "top", "insets": [0.3, 0.2, 0.3, 0.2]},
  {"op": "add_shape", "slide": "$last", "name": "matrix-cell-r1-c3", "shape": "rect", "box": [24.08, 7.35, 8.09, 3.02],
   "fill": "accent1", "line": "none", "text": "**Quick win**\nShip it in the next sprint", "font": {"size": 11, "color": "lt1"},
   "align": "left", "anchor": "top", "insets": [0.3, 0.2, 0.3, 0.2]},
  {"op": "add_shape", "slide": "$last", "name": "matrix-cell-r2-c2", "shape": "rect", "box": [15.68, 10.67, 8.1, 3.03],
   "fill": "accent1+60", "line": "none", "text": "**Plan it**\nWorth it once the data is clean", "font": {"size": 11, "color": "tx1"},
   "align": "left", "anchor": "top", "insets": [0.3, 0.2, 0.3, 0.2]},
  {"op": "add_shape", "slide": "$last", "name": "matrix-cell-r3-c1", "shape": "rect", "box": [7.29, 14, 8.09, 3.02],
   "pattern": {"preset": "wdUpDiag", "fg": "accent6+60", "bg": "lt1"}, "line": "none",
   "text": "**Not sensible**\nHigh effort for little value", "font": {"size": 11, "color": "tx1"},
   "align": "left", "anchor": "top", "insets": [0.3, 0.2, 0.3, 0.2]}
]}
```

A cell that spans two columns takes x of the first and w = right edge of the second − x of the
first (8.09 + 0.3 + 8.1 = 16.49). A badge ("Today") is a small `roundRect` in the highlight color
on the cell's top edge; a path between cells is an `add_line` with `"tail": "triangle"`.

## Process and chevrons

**Use when** the argument is a sequence: steps that build on each other, phases, a flow that ends
in a result. **Not** for parallel items that only look like steps; use tiles.

Structure: a band of chevrons with short step names, one column of explanation under each step
(same columns), optionally a result row. Highlight the step the title is about.

```bash
pptx grid decks/<slug> --cols 4 --area 1.69,4.45,30.48,1.4   # chevron band
pptx grid decks/<slug> --cols 4 --area 1.69,6.15,30.48,5     # explanations, same columns
```

```
1.69,4.45 7.4x1.4 | 9.38,4.45 7.4x1.4 | 17.08,4.45 7.4x1.4 | 24.78,4.45 7.39x1.4

1.69,6.15 7.4x5 | 9.38,6.15 7.4x5 | 17.08,6.15 7.4x5 | 24.78,6.15 7.39x5
```

```json
{"label": "process-onboarding", "ops": [
  {"op": "add_slide", "layout": "Title Only"},
  {"op": "add_shape", "slide": "$last", "name": "step-1", "shape": "homePlate", "box": [1.69, 4.45, 7.4, 1.4],
   "fill": "lt2", "line": "none", "text": "**1 Assess**", "font": {"size": 12, "color": "tx1"},
   "align": "left", "anchor": "middle", "insets": [0.3, 0, 0.6, 0]},
  {"op": "add_shape", "slide": "$last", "name": "step-2", "shape": "chevron", "box": [9.38, 4.45, 7.4, 1.4],
   "fill": "accent1", "line": "none", "text": "**2 Pilot**", "font": {"size": 12, "color": "lt1"},
   "align": "center", "anchor": "middle"},
  {"op": "add_shape", "slide": "$last", "name": "step-2-text", "shape": "rect", "box": [9.38, 6.15, 7.4, 5],
   "fill": "none", "line": "none", "text": "**One team, six weeks:** real work, measured against today",
   "font": {"size": 11, "color": "tx1"}, "align": "left", "anchor": "top", "insets": [0.3, 0.2, 0.3, 0.2]}
]}
```

The first step is a `homePlate` (flat left edge), the following ones are `chevron`s. Keep step
names to one or two words; the explanation goes underneath, never into the chevron.

## Comparison

**Use when** the argument is a choice between alternatives, or before vs. after. **Not** for more
than three options (use a table) or when one side is a straw man.

Structure: one header per option, rows by criterion so the eye compares across, a criterion
column on the left, and a verdict row that states the decision rule (not a repeat of the title).
Emphasize the recommended option; keep the other neutral.

```bash
pptx grid decks/<slug> --cols 2 --gutter 1 --area 7.29,4.45,24.88,1.2                      # option headers
pptx grid decks/<slug> --cols 2 --rows 3 --gutter 1 --row-gutter 0.3 --area 7.29,5.95,24.88,9.17
pptx grid decks/<slug> --cols 1 --rows 3 --area 1.69,5.95,5.3,9.17                         # criteria
```

```
7.29,4.45 11.94x1.2 | 20.23,4.45 11.94x1.2

7.29,5.95 11.94x2.86 | 20.23,5.95 11.94x2.86
7.29,9.11 11.94x2.85 | 20.23,9.11 11.94x2.85
7.29,12.26 11.94x2.86 | 20.23,12.26 11.94x2.86

1.69,5.95 5.3x2.86
1.69,9.11 5.3x2.85
1.69,12.26 5.3x2.86
```

The verdict row takes the full width below: `[1.69, 15.42, 30.48, 1.6]`.

```json
{"label": "comparison-build-buy", "ops": [
  {"op": "add_slide", "layout": "Title Only"},
  {"op": "add_shape", "slide": "$last", "name": "option-buy", "shape": "rect", "box": [20.23, 4.45, 11.94, 1.2],
   "fill": "accent1", "line": "none", "text": "**Buy** (recommended)", "font": {"size": 12, "color": "lt1"},
   "align": "left", "anchor": "middle", "insets": [0.3, 0, 0.3, 0]},
  {"op": "add_shape", "slide": "$last", "name": "criterion-cost", "shape": "rect", "box": [1.69, 5.95, 5.3, 2.86],
   "fill": "none", "line": "none", "text": "**Cost over 3 years**", "font": {"size": 11, "color": "tx1"},
   "align": "left", "anchor": "top", "insets": [0, 0.2, 0.3, 0.2]},
  {"op": "add_shape", "slide": "$last", "name": "buy-cost", "shape": "rect", "box": [20.23, 5.95, 11.94, 2.86],
   "fill": "accent1+60", "line": "none", "text": "**Lower:** licence instead of a team of five",
   "font": {"size": 11, "color": "tx1"}, "align": "left", "anchor": "top", "insets": [0.3, 0.2, 0.3, 0.2]},
  {"op": "add_shape", "slide": "$last", "name": "verdict", "shape": "rect", "box": [1.69, 15.42, 30.48, 1.6],
   "fill": "accent6+80", "line": "none", "text": "**Build only** where the process is what sets us apart",
   "font": {"size": 11, "color": "tx1"}, "align": "left", "anchor": "middle", "insets": [0.3, 0, 0.3, 0]}
]}
```

## Timeline

**Use when** time is the argument: milestones, a roadmap, what happens when. **Not** for a
sequence without dates (use a process).

Structure: dates above a horizontal arrow line, one marker per milestone on the line at each
column's left edge, and under each marker a bold label plus one line. Mark "today" or the critical
milestone with the highlight color.

```bash
pptx grid decks/<slug> --cols 5 --area 1.69,4.45,30.48,1.2   # dates
pptx grid decks/<slug> --cols 5 --area 1.69,6.85,30.48,4     # milestone texts
```

```
1.69,4.45 5.86x1.2 | 7.85,4.45 5.85x1.2 | 14,4.45 5.86x1.2 | 20.16,4.45 5.85x1.2 | 26.31,4.45 5.86x1.2

1.69,6.85 5.86x4 | 7.85,6.85 5.85x4 | 14,6.85 5.86x4 | 20.16,6.85 5.85x4 | 26.31,6.85 5.86x4
```

The line runs at y 6.25 across the content width; markers are 0.4 cm circles centered on it
(y 6.05), at the x of each column. Texts use a left inset of 0 so they start at the marker's edge.

```json
{"label": "timeline-rollout", "ops": [
  {"op": "add_slide", "layout": "Title Only"},
  {"op": "add_line", "slide": "$last", "name": "timeline-axis", "from": [1.69, 6.25], "to": [32.17, 6.25],
   "color": "tx1", "width_pt": 1.5, "tail": "triangle"},
  {"op": "add_shape", "slide": "$last", "name": "milestone-1-date", "shape": "rect", "box": [1.69, 4.45, 5.86, 1.2],
   "fill": "none", "line": "none", "text": "**Q1 2027**", "font": {"size": 12, "color": "accent1"},
   "align": "left", "anchor": "bottom", "insets": [0, 0, 0.3, 0.1]},
  {"op": "add_shape", "slide": "$last", "name": "milestone-1-marker", "shape": "ellipse", "box": [1.69, 6.05, 0.4, 0.4],
   "fill": "accent1", "line": "none"},
  {"op": "add_shape", "slide": "$last", "name": "milestone-1-text", "shape": "rect", "box": [1.69, 6.85, 5.86, 4],
   "fill": "none", "line": "none", "text": "**Pilot live**\nOne region, real customers", "font": {"size": 11, "color": "tx1"},
   "align": "left", "anchor": "top", "insets": [0, 0, 0.3, 0]}
]}
```

## Tiles

**Use when** the slide shows three to six facts of equal weight that together prove the title:
key figures, building blocks, principles. **Not** as the default layout: tiles that all look the
same and argue nothing are an anti-pattern. If one fact matters most, make that tile the
emphasized one.

```bash
pptx grid decks/<slug> --cols 3 --rows 2
```

```
1.69,4.45 9.96x6.14 | 11.95,4.45 9.96x6.14 | 22.21,4.45 9.96x6.14
1.69,10.88 9.96x6.14 | 11.95,10.88 9.96x6.14 | 22.21,10.88 9.96x6.14
```

A key-figure tile: a large number, a bold label, one plain line.

```json
{"label": "tiles-key-figures", "ops": [
  {"op": "add_slide", "layout": "Title Only"},
  {"op": "add_shape", "slide": "$last", "name": "tile-r1-c1", "shape": "rect", "box": [1.69, 4.45, 9.96, 6.14],
   "fill": "lt2", "line": "none", "font": {"color": "tx1"}, "align": "left", "anchor": "top", "insets": [0.4, 0.4, 0.4, 0.4],
   "paragraphs": [
     {"runs": [{"t": "-40 %", "b": true, "size": 28, "color": "accent1"}]},
     {"runs": [{"t": "Lead time", "b": true, "size": 12}]},
     {"runs": [{"t": "from request to release, after six months", "size": 11}]}
   ]}
]}
```

## Chart

**Use when** numbers carry the argument: size (bar or column), trend (line), share of a whole
(doughnut). The title states what the numbers mean, not which numbers they are. One chart per
slide, the series in theme colors, the series that proves the title emphasized.

Structure: the chart over two of three columns, the takeaway (bold lead-in plus one or two lines)
in the third.

```bash
pptx grid decks/<slug> --cols 3
```

```
1.69,4.45 9.96x12.57 | 11.95,4.45 9.96x12.57 | 22.21,4.45 9.96x12.57
```

Chart box: x 1.69, w 21.91 − 1.69 = 20.22; takeaway: the third box.

```json
{"op": "add_chart", "slide": "$last", "name": "chart-lead-time", "box": [1.69, 4.45, 20.22, 12.57],
 "type": "column_clustered", "categories": ["Q1", "Q2", "Q3", "Q4"],
 "series": [{"name": "Lead time (days)", "values": [42, 35, 28, 25]}], "legend": "none", "labels": true}
```

## Table

**Use when** the reader looks values up: many rows of the same attributes, reference data. **Not**
to carry an argument by itself; give it an action title and emphasize the rows that prove it.

Structure: the full content width, or two of three columns with the takeaway in the third (as for
a chart). A header row, few columns, short cell texts. If the brief's building blocks list a
table, clone it and fill it with `set_table` instead; it keeps the exact cell formatting. A table added with
`add_table` already uses the master's default table style; pick another by name from the brief.
For a few values, a comparison or a matrix of text cells reads better.

```bash
pptx grid decks/<slug> --cols 1
```

```
1.69,4.45 30.48x12.57
```

```json
{"op": "add_table", "slide": "$last", "name": "table-options", "box": [1.69, 4.45, 30.48, 6],
 "rows": [["Option", "Cost", "Time to value"], ["Buy", "Low", "3 months"], ["Build", "High", "12 months"]],
 "header": true, "font": {"size": 11}}
```
