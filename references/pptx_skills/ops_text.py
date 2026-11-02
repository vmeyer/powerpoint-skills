"""Text ops: write titles, kickers and body text while keeping formatting the op does not specify."""

import copy
import re

from lxml import etree
from pptx.oxml.ns import qn
from pptx.util import Pt

from .changeset import Context, op, resolve_shape
from .colors import apply_color, parse_color
from .errors import PptxSkillsError
from .slides import resolve_slide

_BULLET_TAGS = ("buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont", "buNone", "buAutoNum", "buChar", "buBlip")
_AFTER_BULLET = ("tabLst", "defRPr", "extLst")


def _shorthand(text: str, font: dict) -> list[dict]:
    """`\\n` separates paragraphs; bold is fully set by the `**` pairs (unless `font.bold` is given)."""
    plain = {} if "bold" in font else {"b": False}
    paragraphs = []
    for line in text.split("\n"):
        runs = [
            {"t": part, "b": True} if i % 2 else {"t": part, **plain}
            for i, part in enumerate(line.split("**"))
            if part
        ]
        paragraphs.append({"runs": runs})
    return paragraphs


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


def _format(run, spec: dict, font: dict, allow_rgb: bool = False) -> None:
    """Apply explicit formatting; the run keeps whatever else it inherited from the template."""
    merged = {**font, **{k: v for k, v in spec.items() if k in ("b", "i", "size", "color")}}
    if "bold" in font and "b" not in spec:
        merged["b"] = font["bold"]
    if "b" in merged:
        run.font.bold = bool(merged["b"])
    if "i" in merged:
        run.font.italic = bool(merged["i"])
    if "size" in merged:
        run.font.size = Pt(merged["size"])
    if "color" in merged:
        ref = parse_color(str(merged["color"]), allow_rgb)
        if ref is None:
            fill = run._r.get_or_add_rPr().find(qn("a:solidFill"))
            if fill is not None:
                fill.getparent().remove(fill)
        else:
            apply_color(run.font.color, ref)


def _is_break(spec: dict) -> bool:
    return bool(spec.get("br"))


def write_paragraphs(text_frame, paragraphs: list[dict], font: dict | None = None, allow_rgb: bool = False) -> None:
    """Replace the text of `text_frame`.

    `paragraphs` is `[{"level"?, "bullet"?, "runs": [{"t", "b"?, "i"?, "color"?, "size"?} | {"br": true}]}]`.
    The first paragraph's properties and first run's properties are the template for everything written, so
    formatting the caller does not specify survives. `font` (size, color, bold) applies to every run.
    Raw `#RRGGBB` colors are rejected unless `allow_rgb`.
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


@op("set_text")
def _set_text(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    shape = resolve_shape(slide, spec["target"])
    if not shape.has_text_frame:
        raise PptxSkillsError(f'shape {shape.shape_id} "{shape.name}" has no text frame')
    paragraphs = _shorthand(spec["text"], spec.get("font") or {}) if "text" in spec else spec["paragraphs"]
    write_paragraphs(shape.text_frame, paragraphs, spec.get("font"))
    ctx.touched.add(slide.slide_id)
