# Scripts

A script is plain [python-pptx](https://python-pptx.readthedocs.io/) code that runs against the working
copy. Use it for what change sets cannot express or would need dozens of ops for: compositions of many
shapes, loops, gradients, freeforms, precise typography, illustrations built from shapes.

```bash
pptx build decks/<slug>/<slug>.pptx decks/<slug>/big-number.py   # run, check, render
pptx exec decks/<slug>/<slug>.pptx decks/<slug>/big-number.py    # run only
```

## Contract

- The script runs inside the tool with the globals below. **It never saves** (`prs.save` raises) and
  touches no other file; the tool saves when the script has finished.
- **All or nothing:** any exception, also a syntax error, leaves the deck as it was; the error names
  the script's line.
- Afterwards the tool gives new slides fresh ids and a section, refreshes slide numbers, runs the
  structure checks (the file must reopen, every relationship must exist, no duplicate ids), keeps a
  snapshot and logs the script as `changes/NNN-<script-name>.py`. Touched are the new slides and those
  whose XML changed; `build` checks and renders exactly those. A change to the slide master, a layout or
  the theme touches every slide.
- `print()` output goes to stderr, so the command's own output (and `--json`) stays clean.
- Paths, e.g. of pictures, are relative to the current folder, the user's project.
- Only the working copy of a deck folder is written (`--outside-deck` allows another `.pptx`, in place).
- Delete or move slides only with the helpers `delete_slide` and `move_slide` (or a change set). A script
  that removes or reorders slides itself is refused, since sections, custom shows and links would fall out
  of step.

| Global | What it is |
|---|---|
| `prs` | the working copy as a python-pptx `Presentation` |
| `deck_dir` | the deck folder as a `Path` (None with `--outside-deck`) |
| `brief` | the style brief as a dict (`theme`, `slide_size_cm`, `content_area`, `layouts`, `text_styles`) or None |
| `cm(x)`, `pt(x)` | lengths for python-pptx |
| `theme_color(target, color)` | colors `shape.fill` (made solid), `shape.line`, `run.font` or any color format such as a gradient stop's `.color`; `color` is a theme role (`accent1`, `accent1+40`, `tx1-25`), `#RRGGBB`, or `none` for a fill or line |
| `plain(shape)` | removes the style python-pptx gives a new auto shape (accent fill, outline, shadow, light centered text), so the shape starts empty; returns it |
| `add_slide(layout, after=None)` | a new slide with the layout's footer, date and number, a fresh id, in the right section; `after` is a slide, a number or `"id:<n>"`. A slide added with `prs.slides.add_slide` alone gets no footer, date or number |
| `delete_slide(slide)`, `move_slide(slide, to)` | remove a slide (with its section entry and the links to it) or move it to the 1-based position `to`; `slide` as for `after` |
| `slide_by_id(id)`, `layout(name)` | a slide by its id, a layout by its name |

Habits that keep scripts short and slides editable:

- Prefer theme colors; they follow `pptx theme` and the master. `#RRGGBB` is fine for a one-off accent.
- Name the shapes (`shape.name = "kpi-value"`) so later change sets can target them.
- Text in a new text box or shape: set `word_wrap`, size, color and alignment yourself.
- `slide.shapes.build_freeform(x, y, scale=cm(1))` puts the pen on the first vertex `(x, y)`;
  `add_line_segments` continues from there, `convert_to_shape(left, top)` places the drawing on the slide.
- Find a shape by name with `next(s for s in slide.shapes if s.name == "house")`. To delete shapes, prefer
  `delete_shape` in a change set, which also removes their animations.
- The slide is `prs.slide_width` × `prs.slide_height` EMU (33.87 × 19.05 cm for 16:9); `brief["content_area"]`
  is `[x, y, w, h]` in cm.

## Example: a big number on a gradient band

```python
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

slide = add_slide("Blank")

band = plain(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, cm(14), prs.slide_height))
band.name = "band"
band.fill.gradient()
band.fill.gradient_angle = 90
theme_color(band.fill.gradient_stops[0].color, "accent1")
theme_color(band.fill.gradient_stops[1].color, "accent1-50")

figure = slide.shapes.add_textbox(cm(1.5), cm(4), cm(11), cm(7))
figure.name = "figure"
frame = figure.text_frame
frame.word_wrap = True
frame.vertical_anchor = MSO_ANCHOR.BOTTOM
run = frame.paragraphs[0].add_run()
run.text = "-50 %"
run.font.size, run.font.bold = pt(96), True
theme_color(run.font, "bg1")
label = frame.add_paragraph().add_run()
label.text = "lead time after the pilot"
label.font.size = pt(20)
theme_color(label.font, "bg1")

claim = slide.shapes.add_textbox(cm(16), cm(5), cm(16), cm(8))
claim.name = "claim"
claim.text_frame.word_wrap = True
for i, (lead, text) in enumerate([("Six weeks,", " one team, real work"), ("Same quality:", " no extra defects")]):
    paragraph = claim.text_frame.paragraphs[0] if i == 0 else claim.text_frame.add_paragraph()
    paragraph.alignment = PP_ALIGN.LEFT
    paragraph.space_after = pt(12)
    for t, bold in ((lead, True), (text, False)):
        r = paragraph.add_run()
        r.text, r.font.bold, r.font.size = t, bold, pt(24)
        theme_color(r.font, "tx1")
```

## Example: a freeform and a row of markers

```python
from pptx.enum.shapes import MSO_SHAPE

slide = prs.slides[-1]

builder = slide.shapes.build_freeform(0, 0, scale=cm(1))  # vertices in cm
builder.add_line_segments([(6, 0), (6, 2.5), (3, 4), (0, 2.5)], close=True)
arrow = plain(builder.convert_to_shape(cm(25), cm(13)))
arrow.name = "pointer"
theme_color(arrow.fill, "accent2")
theme_color(arrow.line, "none")

for i in range(5):
    dot = plain(slide.shapes.add_shape(MSO_SHAPE.OVAL, cm(16 + i * 1.2), cm(15), cm(0.6), cm(0.6)))
    dot.name = f"marker-{i + 1}"
    theme_color(dot.fill, "accent1" if i < 3 else "tx1+75")
```
