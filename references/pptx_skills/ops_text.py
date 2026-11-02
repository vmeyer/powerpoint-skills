"""Text ops: write titles, kickers and body text while keeping formatting the op does not specify."""

import copy
import re

from lxml import etree
from pptx.oxml.ns import qn
from pptx.util import Pt

from .changeset import Context, op, resolve_shape
from .colors import RGB_OPTION, apply_color, parse_color
from . import timing
from .errors import PptxSkillsError
from .slides import resolve_slide

_BULLET_TAGS = ("buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont", "buNone", "buAutoNum", "buChar", "buBlip")
_AFTER_BULLET = ("tabLst", "defRPr", "extLst")


def _shorthand(text: str, font: dict) -> list[dict]:
    """The `text` shorthand as paragraphs: `\\n` separates paragraphs, `**…**` marks bold runs (R13′).

    With markers every weight is written: marked runs bold, the others not (or `font.bold`). Without markers
    no weight is written and a bold copied from the existing first run is dropped (`"b": None`), so the weight
    comes from the layout or master, unless `font.bold` sets it. Shared by text, shape and table ops.
    """
    plain = False if "**" in text and "bold" not in font else None
    return [
        {"runs": [{"t": part, "b": True if i % 2 else plain} for i, part in enumerate(line.split("**")) if part]}
        for line in text.split("\n")
    ]


def _drop_bullets(ppr) -> None:
    for tag in _BULLET_TAGS:
        for el in ppr.findall(qn(f"a:{tag}")):
            ppr.remove(el)


def _style_paragraph(p_el, spec: dict) -> None:
    """Apply level and bullet on top of the template pPr; stale margins never survive a change."""
    level, bullet = spec.get("level"), spec.get("bullet")
    if level is None and bullet is None:
        return  # the template's level, bullet and margins carry over as they are
    ppr = p_el.get_or_add_pPr()
    if level is not None:
        ppr.set("lvl", str(level))
        ppr.attrib.pop("marL", None)
        ppr.attrib.pop("indent", None)
    if bullet is False:
        _drop_bullets(ppr)
        ppr.set("marL", "0")
        ppr.set("indent", "0")
        none = etree.SubElement(ppr, qn("a:buNone"))
        anchor = next((c for c in ppr if c.tag in {qn(f"a:{t}") for t in _AFTER_BULLET}), None)
        if anchor is not None:
            anchor.addprevious(none)
    elif bullet:
        has_bullet = any(ppr.find(qn(f"a:{t}")) is not None for t in ("buChar", "buAutoNum", "buBlip"))
        if not has_bullet:
            # let the layout or master supply the bullet and its margins
            for el in ppr.findall(qn("a:buNone")):
                ppr.remove(el)
            ppr.attrib.pop("marL", None)
            ppr.attrib.pop("indent", None)


def _format(run, spec: dict, font: dict, allow_rgb: bool | None = None) -> None:
    """Apply explicit formatting; the run keeps whatever else it inherited from the template.

    A run's own `b` wins over `font.bold`; `"b": None` (the shorthand without markers) removes the template's
    weight unless `font.bold` is given. `allow_rgb` None means the op has no allow_rgb option.
    """
    if spec.get("b") is not None:
        run.font.bold = bool(spec["b"])
    elif "bold" in font:
        run.font.bold = bool(font["bold"])
    elif "b" in spec:
        run.font.bold = None  # inherit the weight from the layout or master
    merged = {**font, **{k: v for k, v in spec.items() if k in ("i", "size", "color")}}
    if "i" in merged:
        run.font.italic = bool(merged["i"])
    if "size" in merged:
        run.font.size = Pt(merged["size"])
    if "color" in merged:
        ref = parse_color(str(merged["color"]), bool(allow_rgb), RGB_OPTION if allow_rgb is not None else None)
        if ref is None:
            fill = run._r.get_or_add_rPr().find(qn("a:solidFill"))
            if fill is not None:
                fill.getparent().remove(fill)
        else:
            apply_color(run.font.color, ref)


def _is_break(spec: dict) -> bool:
    return bool(spec.get("br"))


def write_paragraphs(text_frame, paragraphs: list[dict], font: dict | None = None, allow_rgb: bool | None = None) -> None:
    """Replace the text of `text_frame`.

    `paragraphs` is `[{"level"?, "bullet"?, "runs": [{"t", "b"?, "i"?, "color"?, "size"?} | {"br": true}]}]`.
    The first paragraph's properties and first run's properties are the template for everything written, so
    formatting the caller does not specify survives. `font` (size, color, bold) applies to every run.
    Raw `#RRGGBB` colors are rejected unless `allow_rgb`; None means the op offers no allow_rgb option.
    """
    font = font or {}
    body = text_frame._txBody
    first = body.findall(qn("a:p"))[0] if body.findall(qn("a:p")) else body.add_p()
    first_run = first.find(qn("a:r"))
    template_rpr = None
    if first_run is not None and first_run.find(qn("a:rPr")) is not None:
        template_rpr = copy.deepcopy(first_run.find(qn("a:rPr")))
    for p_el in body.findall(qn("a:p")):
        if p_el is not first:
            body.remove(p_el)
    for child in list(first):
        if child.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            first.remove(child)
    blank = copy.deepcopy(first)  # pPr and endParaRPr only
    first.getparent().remove(first)
    for spec in paragraphs or [{"runs": []}]:
        p_el = copy.deepcopy(blank)
        body.append(p_el)
        _style_paragraph(p_el, spec)
        para = text_frame.paragraphs[-1]
        for run_spec in spec.get("runs", []):
            if _is_break(run_spec):
                para.add_line_break()
                continue
            run = para.add_run()
            run.text = run_spec["t"]
            if template_rpr is not None:
                old = run._r.find(qn("a:rPr"))
                if old is not None:
                    run._r.remove(old)
                run._r.insert(0, copy.deepcopy(template_rpr))
            _format(run, run_spec, font, allow_rgb)


def slide_number(ctx: Context, slide) -> int:
    return [s.slide_id for s in ctx.prs.slides].index(slide.slide_id) + 1


def prune_text_animation(ctx: Context, slide, shape) -> None:
    """After new text: drop the animation steps of paragraphs the shape no longer has (R27)."""
    count = len(shape.text_frame.paragraphs)
    if timing.drop_paragraphs(slide._element, shape.shape_id, count):
        ctx.messages.append(
            f'slide {slide_number(ctx, slide)}: "{shape.name}" now has {count} paragraph(s); the animation steps of '
            "the paragraphs that are gone were removed"
        )


@op("set_text")
def _set_text(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    shape = resolve_shape(slide, spec["target"])
    if not shape.has_text_frame:
        raise PptxSkillsError(f'shape {shape.shape_id} "{shape.name}" has no text frame')
    paragraphs = _shorthand(spec["text"], spec.get("font") or {}) if "text" in spec else spec["paragraphs"]
    write_paragraphs(shape.text_frame, paragraphs, spec.get("font"))
    prune_text_animation(ctx, slide, shape)
    ctx.touched.add(slide.slide_id)
