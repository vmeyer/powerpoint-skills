# Review checklist

Go through this list every round, looking at the rendered PNG with your image-viewing tool
(`decks/<slug>/render/slide-NN.png`), after `pptx check` has no errors left. Fix what fails with
the next change set. At most three rounds; then show the slide and list what is still open.

## Message

- [ ] The title states the takeaway as a full sentence (or two short ones) in at most two lines;
      it is not a topic ("Overview", "Market").
- [ ] The kicker is a short label for the slide's role or chapter, not a sentence.
- [ ] Everything on the slide supports the title; nothing repeats it (no "key takeaway" box).

## Structure

- [ ] The eye follows the argument: where it starts, where it goes, where it ends.
- [ ] The visual form fits the argument (matrix, process, comparison, timeline, tiles, chart, table).
- [ ] The hierarchy is clear: title, then structure labels, then explanations. Not everything is an
      identical card.

## Text

- [ ] No keyword chains joined by " · ", " | " or " • ".
- [ ] Little text: a bold label plus one short line, bold lead-in plus plain text; no paragraph
      longer than three lines, no bullet wall. Details belong in the speaker notes.
- [ ] No awkward line breaks: no single word alone on the last line, no compound word broken
      across lines, no label wrapping where a slightly wider box or shorter word would fit.
- [ ] Concrete examples rather than buzzwords.
- [ ] Few sizes (title plus one or two body sizes), none below the brief's minimum.

## Alignment and space

- [ ] Edges line up; gutters are equal; equal roles have equal sizes.
- [ ] Nothing is cramped: text keeps its inset from the shape edge, shapes keep the margin to the
      content-area edges, the footer zone stays free.
- [ ] No text is clipped or runs past its shape (the render can wrap slightly differently from
      PowerPoint, so a box filled to its last line needs more room).

## Color and contrast

- [ ] Only theme colors; one family for value and emphasis, neutral fills for structure, the
      highlight color at most once or twice.
- [ ] Readable contrast: light text only on strong fills, dark text on light tints and hatching.
- [ ] No decoration without meaning: no icons, gradients or shadows the master does not use.

## Deck and feedback

- [ ] Consistent with the neighbouring slides: title position, kicker use, footer, the same meaning
      in the same color.
- [ ] The user's feedback is actually addressed, every point of it, and nothing else changed.
- [ ] If the input was a screenshot: compare the render with it side by side; the slide in the
      screenshot is the slide you changed, and the change shows where the user marked it.

## What to report when you show the slide

1. The PNG path.
2. One line on what the slide argues and how the structure shows it.
3. Open findings: `check` warnings you kept and why, review points still open after three rounds,
   and any theme font that was missing (the render used a fallback).
