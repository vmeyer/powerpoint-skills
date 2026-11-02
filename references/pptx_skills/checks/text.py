"""Text rules: overflow, leftover template text, sizes, line and bullet counts, keyword chains."""

import re

from pptx.oxml.ns import qn

from ..fonts import _Inheritance, _layout_placeholder, effective_run_style, measure_paragraphs
from ..geometry import absolute_box
from ..inventory import plain_text, table_cell_texts
from ..units import EMU_PER_CM
from . import rule

OVERFLOW_TOLERANCE_CM = 0.05
_SMALL_ROLES = ("footer", "number", "date")
_LEFTOVER = re.compile(r"\bLorem\b|\bTODO\b|\bXXX\b")
_CHAIN_SEPARATOR = re.compile(r"[·•|]|\s/\s")
_BULLET_TAGS = {qn("a:buNone"): False, qn("a:buChar"): True, qn("a:buAutoNum"): True, qn("a:buBlip"): True}


def _norm(text: str) -> str:
    return " ".join(text.split()).casefold()


def _paragraph_texts(element) -> list[str]:
    body = element.find(qn("p:txBody"))
    if body is None:
        return []
    return ["".join(t.text or "" for t in p.iter(qn("a:t"))) for p in body.findall(qn("a:p"))]


def _text_shapes(slide):
    """(inventory dict, shape) of shapes that carry text."""
    for item, shape in slide.shapes():
        if "".join(_paragraph_texts(shape._element)).strip():
            yield item, shape


def _scale(measure):
    return measure.font_scale if measure.autofit == "shrink" else 1.0


@rule("overflow", "text", "error")
def overflow(ctx, slide):
    out = []
    for item, shape in _text_shapes(slide):
        m = measure_paragraphs(shape, slide.pptx)
        if m.autofit == "shape":
            continue
        scale = _scale(m)
        reduction = m.line_spacing_reduction if m.autofit == "shrink" else 0.0
        height = m.height_cm * scale * (1 - reduction)
        box = absolute_box(shape)
        if height > box.h + OVERFLOW_TOLERANCE_CM:
            out.append(ctx.finding(slide, item, f"the text needs {height:.2f} cm but the box is {box.h:.2f} cm high", "shorten the text, enlarge the box or split the content over two slides"))
            continue
        insets = _Inheritance(shape, slide.pptx).body().insets
        room = box.w - (insets["lIns"] + insets["rIns"]) / EMU_PER_CM
        if m.widest_word_cm * scale > room + 0.01:
            out.append(ctx.finding(slide, item, f"a word is {m.widest_word_cm * scale:.2f} cm wide but the box leaves {room:.2f} cm", "widen the box or use a shorter word"))
    return out


def _is_prompt(text: str, prompt: str) -> bool:
    a, b = _norm(text), _norm(prompt)
    if not a or not b:
        return False
    if a.startswith(b):
        return True
    # slide text that is only the beginning of the prompt counts when it is a whole-word, substantial prefix
    return b.startswith(a) and b[len(a)] == " " and (len(a.split()) >= 2 or len(a) >= 12)


@rule("leftover-text", "text", "error")
def leftover_text(ctx, slide):
    out = []
    cloned = (ctx.deck.get("cloned_texts") or {}).get(str(slide.id), {})
    for item, shape in _text_shapes(slide):
        text = "\n".join(_paragraph_texts(shape._element))
        why = None
        ph = shape._element.ph
        layout_ph = _layout_placeholder(ph, slide.pptx.slide_layout) if ph is not None and item["role"] not in _SMALL_ROLES else None
        if layout_ph is not None and _is_prompt(text, "\n".join(_paragraph_texts(layout_ph))):
            why = "it is still the layout's prompt text"
        elif (match := _LEFTOVER.search(text)):
            why = f"it contains {match.group(0)!r}"
        elif str(item["id"]) in cloned and _norm(cloned[str(item["id"])]) == _norm(plain_text(item["text"])):
            why = "it is still the text of the cloned example slide"
        if why:
            out.append(ctx.finding(slide, item, f"template text is left over: {why}", "replace the template text with the real content"))
    for item, shape in slide.shapes():
        if item["kind"] == "table":
            out += _leftover_cells(ctx, slide, item, shape, cloned)
    return out


def _leftover_cells(ctx, slide, item, shape, cloned) -> list:
    """Table cells that still hold Lorem/TODO/XXX or the text of the cloned example table."""
    out = []
    for r, row in enumerate(table_cell_texts(shape)):
        for c, text in enumerate(row):
            if (match := _LEFTOVER.search(text)):
                why = f"it contains {match.group(0)!r}"
            elif (key := f"{item['id']}:r{r}c{c}") in cloned and _norm(cloned[key]) == _norm(text):
                why = "it is still the text of the cloned example slide"
            else:
                continue
            out.append(ctx.finding(slide, item, f"template text is left over in table cell row {r + 1}, column {c + 1}: {why}", "replace the template text with the real content"))
    return out


def _limit_for(ctx, slide, item, box) -> float:
    area = ctx.content_area
    small = item["role"] in _SMALL_ROLES or (area is not None and box.y >= area[1] + area[3])
    return ctx.rules["min_size_small_pt" if small else "min_size_pt"]


@rule("min-size", "text", "error")
def min_size(ctx, slide):
    out = []
    for item, shape in _text_shapes(slide):
        limit = _limit_for(ctx, slide, item, absolute_box(shape))
        m = measure_paragraphs(shape, slide.pptx)
        scale = _scale(m)
        smallest = None
        for p in shape._element.find(qn("p:txBody")).findall(qn("a:p")):
            for run in p:
                if run.tag in (qn("a:r"), qn("a:fld")) and (run.findtext(qn("a:t")) or "").strip():
                    size = effective_run_style(run, p, shape, slide.pptx).size_pt * scale
                    smallest = size if smallest is None else min(smallest, size)
        if smallest is not None and smallest < limit - 1e-6:
            out.append(ctx.finding(slide, item, f"text is shown at {smallest:g} pt, below the minimum of {limit:g} pt", f"raise the font size to at least {limit:g} pt or cut text"))
    return out


@rule("too-many-lines", "text", "warning")
def too_many_lines(ctx, slide):
    out = []
    for item, shape in _text_shapes(slide):
        lines = measure_paragraphs(shape, slide.pptx).lines_per_paragraph
        if item["role"] == "title":
            if sum(lines) > ctx.rules["max_title_lines"]:
                out.append(ctx.finding(slide, item, f"the title wraps to {sum(lines)} lines (max {ctx.rules['max_title_lines']})", f"shorten the title to ≤ {ctx.rules['max_title_lines']} lines"))
        elif max(lines, default=0) > ctx.rules["max_lines"]:
            out.append(ctx.finding(slide, item, f"a paragraph wraps to {max(lines)} lines (max {ctx.rules['max_lines']})", f"shorten to ≤ {ctx.rules['max_lines']} lines or widen the box"))
    return out


def _is_bullet(paragraph, inheritance) -> bool:
    p_pr = paragraph.find(qn("a:pPr"))
    sources = ([p_pr] if p_pr is not None else []) + inheritance.levels(paragraph)
    for source in sources:
        for child in source:
            if child.tag in _BULLET_TAGS:
                return _BULLET_TAGS[child.tag]
    return False


@rule("too-many-bullets", "text", "warning")
def too_many_bullets(ctx, slide):
    out = []
    for item, shape in _text_shapes(slide):
        inheritance = _Inheritance(shape, slide.pptx)
        paragraphs = shape._element.find(qn("p:txBody")).findall(qn("a:p"))
        count = sum(1 for p in paragraphs if "".join(t.text or "" for t in p.iter(qn("a:t"))).strip() and _is_bullet(p, inheritance))
        if count > ctx.rules["max_bullets"]:
            out.append(ctx.finding(slide, item, f"{count} bullets in one box (max {ctx.rules['max_bullets']})", "cut to the strongest points or split into columns or two slides"))
    return out


@rule("keyword-chain", "text", "warning")
def keyword_chain(ctx, slide):
    out = []
    for item, shape in _text_shapes(slide):
        for text in _paragraph_texts(shape._element):
            parts = [p.strip() for p in _CHAIN_SEPARATOR.split(text) if p.strip()]
            if len(parts) >= 3 and all(len(p.split()) <= 4 for p in parts):
                out.append(ctx.finding(slide, item, f"keywords are chained with separators: {text.strip()!r}", "write a sentence or use one line per item"))
                break
    return out


@rule("shrink-to-fit", "text", "warning")
def shrink_to_fit(ctx, slide):
    out = []
    for item, shape in _text_shapes(slide):
        m = measure_paragraphs(shape, slide.pptx)
        if m.autofit == "shrink" and m.font_scale < 1:
            out.append(ctx.finding(slide, item, f"PowerPoint shrinks the text to {m.font_scale:.0%}", "shorten the text or enlarge the box so no shrinking is needed"))
    return out


@rule("empty-placeholder", "text", "warning")
def empty_placeholder(ctx, slide):
    out = []
    for item, shape in slide.shapes():
        if item["kind"] != "placeholder" or item["role"] in _SMALL_ROLES or shape._element.tag != qn("p:sp"):
            continue
        if not "".join(_paragraph_texts(shape._element)).strip():
            out.append(ctx.finding(slide, item, "the placeholder is empty", "fill it or delete the placeholder"))
    return out


@rule("topic-title", "text", "info")
def topic_title(ctx, slide):
    out = []
    for item, shape in _text_shapes(slide):
        words = len(" ".join(_paragraph_texts(shape._element)).split())
        if item["role"] == "title" and words < ctx.rules["title_min_words"]:
            out.append(ctx.finding(slide, item, f"the title has {words} words, which reads as a topic", "write the takeaway as a full sentence"))
    return out
