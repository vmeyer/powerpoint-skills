# Visual forms: ideas, not templates

Start from what the slide argues, then pick the form that makes the argument visible before anyone
reads a word. The forms below are starting points; combine them, break them, invent your own.

| The argument is about … | Forms that show it |
|---|---|
| one striking fact | a big number, a full-bleed color block, a single quote, one picture with a short caption |
| a contrast | before / after, two columns of unequal weight, an asymmetric split with the winner larger |
| a sequence | a process of chevrons, a path of numbered markers, a staircase, a timeline |
| proportion or trend | a chart with the decisive series in the accent color and the rest in grey |
| two dimensions | a matrix whose cells are colored by meaning, a 2 × 2 with the target quadrant highlighted |
| parts of a whole | a stacked bar, a doughnut, shapes sized by share |
| a few parallel facts | tiles with one emphasized, icons drawn from shapes, a row of big figures |
| many values to look up | a table with the rows that prove the title emphasized |

If no form fits, the slide may carry two messages: split it.

## Bold moves

- **Scale:** one number at 80–120 pt carries a slide; the label beneath it at body size.
- **Color as signal:** one accent for what matters, grey or tints for the rest. A full-bleed accent
  band or background on a key slide breaks the rhythm of a deck on purpose.
- **Asymmetry:** a 1/3 – 2/3 split, a picture bleeding off the edge, a statement set low and large.
- **Whitespace:** fewer, larger elements. Empty space makes the remaining shapes count.
- **Drawn illustration:** arrows, brackets, markers, freeforms and gradients from shapes (see
  [scripts.md](scripts.md)) instead of clip art; the user's photos and screenshots where they exist.
- **Rhythm across a deck:** alternate dense and sparse slides, light and dark ones; keep title
  position, fonts and the meaning of each color the same.

## Sketches

A big-number slide as a change set (the title states what the number means):

```json
{"label": "big-number", "ops": [
  {"op": "add_slide", "layout": "Title Only"},
  {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "Half of all requests now ship within a week"},
  {"op": "add_shape", "slide": "$last", "shape": "rect", "box": [0, 4.5, 14, 14.55], "fill": "accent1",
   "paragraphs": [{"runs": [{"t": "52 %", "b": true, "size": 96}]}, {"runs": [{"t": "ship within 7 days"}]}],
   "font": {"color": "bg1"}, "anchor": "bottom", "insets": [1.7, 0, 1, 1.5]},
  {"op": "add_shape", "slide": "$last", "shape": "rect", "box": [16, 6, 16, 8],
   "text": "**A year ago:** 12 %\n**Driver:** small batches and automated tests", "font": {"size": 20}}
]}
```

A before / after contrast where the winner takes more room:

```json
{"label": "before-after", "ops": [
  {"op": "add_slide", "layout": "Title Only"},
  {"op": "set_text", "slide": "$last", "target": {"role": "title"}, "text": "Small batches replaced the quarterly big bang"},
  {"op": "add_shape", "slide": "$last", "shape": "rect", "box": [1.7, 5, 10, 11], "fill": "tx1+85",
   "text": "**Before**\nOne release per quarter, two weeks of testing", "font": {"color": "tx1+25"}},
  {"op": "add_shape", "slide": "$last", "shape": "rightArrow", "box": [12.2, 9.5, 2, 2], "fill": "accent2"},
  {"op": "add_shape", "slide": "$last", "shape": "rect", "box": [14.7, 5, 17.5, 11], "fill": "accent1",
   "text": "**After**\nA release every day, tests run in minutes", "font": {"size": 28, "color": "bg1"}}
]}
```

Process, timeline, matrix and tiles are built from the same ops: `chevron` and `homePlate` for steps,
`add_line` with markers for time, filled `rect`s whose color carries meaning for a matrix, `pptx grid`
for boxes that must line up. When one form repeats on many slides, give each slide a twist: a different
emphasis, scale or color.
