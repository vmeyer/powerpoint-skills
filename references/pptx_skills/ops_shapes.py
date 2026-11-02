"""Shape ops: build slides from native shapes with theme colors, snapped to the master's grid."""

from pptx.enum.dml import MSO_LINE, MSO_PATTERN_TYPE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt
from lxml import etree

from . import timing
from .changeset import Context, _walk, op, resolve_shape
from .colors import apply_color, parse_color
from .errors import PptxSkillsError
from .geometry import Box, _xfrm, snap_box
from .ops_text import _format, _shorthand, prune_text_animation, slide_number, write_paragraphs
from .slides import resolve_slide

_PRESETS = {
    "rect": MSO_SHAPE.RECTANGLE,
    "roundRect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "homePlate": MSO_SHAPE.PENTAGON,
    "chevron": MSO_SHAPE.CHEVRON,
    "rightArrow": MSO_SHAPE.RIGHT_ARROW,
    "leftArrow": MSO_SHAPE.LEFT_ARROW,
    "upArrow": MSO_SHAPE.UP_ARROW,
    "downArrow": MSO_SHAPE.DOWN_ARROW,
    "ellipse": MSO_SHAPE.OVAL,
}
_ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
_ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}
_STYLE_FIELDS = ("fill", "line", "pattern", "text", "paragraphs", "font", "align", "anchor", "insets", "radius")


def _snap_lines(ctx: Context) -> tuple[list[float], list[float], float] | None:
    """Content-area edges plus brief guides, and the tolerance; None without a brief."""
    brief = ctx.brief
    if not brief:
        return None
    xs, ys = list(brief.get("guides", {}).get("vertical", [])), list(brief.get("guides", {}).get("horizontal", []))
    area = brief.get("content_area")
    if area:
        xs += [area[0], area[0] + area[2]]
        ys += [area[1], area[1] + area[3]]
    return xs, ys, float(brief.get("rules", {}).get("snap_cm", 0.05))


def _snapped(ctx: Context, box: list[float]) -> Box:
    raw = Box(*box)
    lines = _snap_lines(ctx)
    return snap_box(raw, lines[0], lines[1], lines[2]) if lines else raw


def _snap_point(ctx: Context, x: float, y: float) -> tuple[float, float]:
    lines = _snap_lines(ctx)
    if not lines:
        return x, y
    box = snap_box(Box(x, y, 0, 0), lines[0], lines[1], lines[2])
    return box.x, box.y


def _check_name(slide, name: str) -> None:
    if any(s.name == name for s in _walk(slide.shapes)):
        raise PptxSkillsError(f'a shape named "{name}" already exists on the slide')


def shape_name(ctx: Context, slide, spec: dict, prefix: str) -> str:
    """The op's `name`, which must be new on the slide, or else the first free `<prefix>-<n>`.

    A generated name is recorded in `ctx.auto_names` so the output can report it.
    """
    if "name" in spec:
        _check_name(slide, spec["name"])
        return spec["name"]
    taken = {s.name for s in _walk(slide.shapes)}
    n = 1
    while f"{prefix}-{n}" in taken:
        n += 1
    ctx.auto_names.setdefault(slide.slide_id, []).append(f"{prefix}-{n}")
    return f"{prefix}-{n}"


def _color(text: str, what: str):
    try:
        return parse_color(str(text))
    except PptxSkillsError as exc:
        raise PptxSkillsError(f"{what}: {exc}") from exc


def _set_fill(shape, spec: dict) -> None:
    if "fill" in spec and "pattern" in spec:
        raise PptxSkillsError('"fill" and "pattern" are mutually exclusive')
    if "pattern" in spec:
        pat = spec["pattern"]
        try:
            kind = MSO_PATTERN_TYPE.from_xml(pat["preset"])
        except (KeyError, ValueError) as exc:
            raise PptxSkillsError(f'unknown pattern preset "{pat["preset"]}" (OOXML names such as wdUpDiag, ltUpDiag, dkDnDiag)') from exc
        shape.fill.patterned()
        shape.fill.pattern = kind
        for key, target in (("fg", shape.fill.fore_color), ("bg", shape.fill.back_color)):
            ref = _color(pat[key], f"pattern {key}")
            if ref is None:
                raise PptxSkillsError(f'pattern {key} needs a color, not "none"')
            apply_color(target, ref)
    elif "fill" in spec:
        ref = _color(spec["fill"], "fill")
        if ref is None:
            shape.fill.background()
        else:
            shape.fill.solid()
            apply_color(shape.fill.fore_color, ref)


def _set_line(shape, line) -> None:
    ln = shape.line
    if line == "none":
        ln.fill.background()
        return
    ref = _color(line["color"], "line color")
    if ref is None:
        ln.fill.background()
        return
    ln.fill.solid()
    apply_color(ln.fill.fore_color, ref)
    if "width_pt" in line:
        ln.width = Pt(line["width_pt"])
    if "dash" in line:
        try:
            ln.dash_style = MSO_LINE.from_xml(line["dash"])
        except ValueError as exc:
            raise PptxSkillsError(f'unknown dash "{line["dash"]}" (solid, dash, sysDash, dot, sysDot, lgDash, dashDot, lgDashDot, lgDashDotDot)') from exc


def _set_text(shape, spec: dict, wrap: bool = False) -> None:
    tf = shape.text_frame
    font = spec.get("font") or {}
    if "text" in spec or "paragraphs" in spec:
        paragraphs = _shorthand(spec["text"], font) if "text" in spec else spec["paragraphs"]
        write_paragraphs(tf, paragraphs, font)
        wrap = True
    elif font:  # font alone restyles the existing runs
        for para in tf.paragraphs:
            for run in para.runs:
                _format(run, {}, font)
    if wrap:
        tf.word_wrap = True
    if "align" in spec:
        for para in tf.paragraphs:
            para.alignment = _ALIGN[spec["align"]]
    if "anchor" in spec:
        tf.vertical_anchor = _ANCHOR[spec["anchor"]]
    if "insets" in spec:
        l, t, r, b = spec["insets"]
        tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = Cm(l), Cm(t), Cm(r), Cm(b)


def _style(shape, spec: dict, wrap: bool = False) -> None:
    _set_fill(shape, spec)
    if "line" in spec:
        _set_line(shape, spec["line"])
    if "radius" in spec:
        if len(shape.adjustments) == 0 or shape.auto_shape_type != MSO_SHAPE.ROUNDED_RECTANGLE:
            raise PptxSkillsError('"radius" only applies to roundRect shapes')
        shape.adjustments[0] = spec["radius"]
    if shape.has_text_frame:
        _set_text(shape, spec, wrap)


def _place(shape, box: Box) -> None:
    shape.left, shape.top, shape.width, shape.height = Cm(box.x), Cm(box.y), Cm(box.w), Cm(box.h)


TEXT_INSETS_CM = [0.2, 0.2, 0.2, 0.2]


def _text_defaults(ctx: Context, spec: dict) -> dict:
    """`spec` with what an added shape's text gets unless the op says otherwise: the deck's default text size
    (`ctx.text_size`) in tx1, left-aligned at the top, 0.2 cm from the edges. Without text nothing is added."""
    if "text" not in spec and "paragraphs" not in spec:
        return spec
    font = {"size": ctx.text_size, "color": "tx1", **(spec.get("font") or {})}
    return {"align": "left", "anchor": "top", "insets": TEXT_INSETS_CM, **spec, "font": font}


@op("add_shape")
def _add_shape(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    kind = spec["shape"]
    name = shape_name(ctx, slide, spec, kind)
    box = _snapped(ctx, spec["box"])
    if kind == "textbox":
        shape = slide.shapes.add_textbox(Cm(box.x), Cm(box.y), Cm(box.w), Cm(box.h))
    else:
        shape = slide.shapes.add_shape(_PRESETS[kind], Cm(box.x), Cm(box.y), Cm(box.w), Cm(box.h))
    shape.name = name
    spec = _text_defaults(ctx, spec)
    if "rotation" in spec:
        shape.rotation = spec["rotation"]
    style = shape._element.find(qn("p:style"))
    if style is not None:  # unstyled shapes have no fill, line or effects and inherit the default text style
        shape._element.remove(style)
    _style(shape, spec, wrap=True)
    ctx.touched.add(slide.slide_id)


@op("add_line")
def _add_line(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    name = shape_name(ctx, slide, spec, "line")
    (x1, y1), (x2, y2) = (_snap_point(ctx, *spec[k]) for k in ("from", "to"))
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    line.name = name
    style = {"color": spec["color"], "width_pt": spec["width_pt"]}
    if "dash" in spec:
        style["dash"] = spec["dash"]
    _set_line(line, style)
    ln = line.line._get_or_add_ln()
    for key, tag in (("head", "a:headEnd"), ("tail", "a:tailEnd")):
        if key in spec:
            etree.SubElement(ln, qn(tag)).set("type", spec[key])
    ctx.touched.add(slide.slide_id)


def _is_connector(shape) -> bool:
    return shape._element.tag == qn("p:cxnSp")


@op("update_shape")
def _update_shape(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    shape = resolve_shape(slide, spec["target"])
    if _is_connector(shape) and set(spec) & {"box", "rotation", "fill", "pattern", "text", "paragraphs", "radius", "align", "anchor", "insets"}:
        raise PptxSkillsError("a line only takes name and line; delete it and add a new line to move it")
    if "name" in spec and spec["name"] != shape.name:
        _check_name(slide, spec["name"])
        shape.name = spec["name"]
    if "box" in spec:
        if shape._element.getparent().tag == qn("p:grpSp"):
            raise PptxSkillsError(f'shape "{shape.name}" is inside a group; ungroup it before changing its box')
        _place(shape, _snapped(ctx, spec["box"]))
    if "rotation" in spec:
        shape.rotation = spec["rotation"]
    if (set(spec) & {"text", "paragraphs", "align", "anchor", "insets"}) and not shape.has_text_frame:
        raise PptxSkillsError(f'shape {shape.shape_id} "{shape.name}" has no text frame')
    if shape.shape_type is not None and shape.shape_type.name == "GROUP" and set(spec) & set(_STYLE_FIELDS):
        raise PptxSkillsError(f'"{shape.name}" is a group; change its children')
    _style(shape, spec)
    if "text" in spec or "paragraphs" in spec:
        prune_text_animation(ctx, slide, shape)
    ctx.touched.add(slide.slide_id)


@op("delete_shape")
def _delete_shape(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    shape = resolve_shape(slide, spec["target"])
    element = shape._element
    element.getparent().remove(element)
    if timing.drop_shapes(slide._element, timing.shape_ids(element)):
        ctx.messages.append(f'slide {slide_number(ctx, slide)}: the animation of "{shape.name}" was removed with it')
    ctx.touched.add(slide.slide_id)


_SHAPE_TAGS = {qn(f"p:{t}") for t in ("sp", "grpSp", "graphicFrame", "cxnSp", "pic", "contentPart")}


@op("z_order")
def _z_order(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    element = resolve_shape(slide, spec["target"])._element
    parent = element.getparent()
    siblings = [c for c in parent if c.tag in _SHAPE_TAGS]
    index, to = siblings.index(element), spec["to"]
    if to == "front":
        siblings[-1].addnext(element)
    elif to == "back":
        siblings[0].addprevious(element)
    elif to == "forward" and index < len(siblings) - 1:
        siblings[index + 1].addnext(element)
    elif to == "backward" and index > 0:
        siblings[index - 1].addprevious(element)
    ctx.touched.add(slide.slide_id)


@op("group")
def _group(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    shapes = [resolve_shape(slide, t) for t in spec["targets"]]
    if len({s.shape_id for s in shapes}) != len(shapes):
        raise PptxSkillsError("a shape is listed twice in targets")
    for s in shapes:
        if s._element.getparent() is not slide.shapes._spTree:
            raise PptxSkillsError(f'shape "{s.name}" is inside a group; only top-level shapes can be grouped')
        if s.is_placeholder:
            raise PptxSkillsError(f'shape "{s.name}" is a placeholder and cannot be grouped')
    name = shape_name(ctx, slide, spec, "group")
    group = slide.shapes.add_group_shape(shapes)  # extents are recalculated so chOff/chExt equal off/ext
    group.name = name
    ctx.touched.add(slide.slide_id)


def _convert(xfrm, group_xfrm) -> None:
    off, ext = group_xfrm.find(qn("a:off")), group_xfrm.find(qn("a:ext"))
    ch_off, ch_ext = group_xfrm.find(qn("a:chOff")), group_xfrm.find(qn("a:chExt"))
    gx, gy, gw, gh = (int(v) for v in (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy")))
    cx, cy, cw, ch = (int(v) for v in (ch_off.get("x"), ch_off.get("y"), ch_ext.get("cx"), ch_ext.get("cy")))
    sx, sy = (gw / cw if cw else 1.0), (gh / ch if ch else 1.0)
    c_off, c_ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    if c_off is None or c_ext is None:
        return
    c_off.set("x", str(round(gx + (int(c_off.get("x")) - cx) * sx)))
    c_off.set("y", str(round(gy + (int(c_off.get("y")) - cy) * sy)))
    c_ext.set("cx", str(round(int(c_ext.get("cx")) * sx)))
    c_ext.set("cy", str(round(int(c_ext.get("cy")) * sy)))


@op("ungroup")
def _ungroup(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    shape = resolve_shape(slide, spec["target"])
    if shape._element.tag != qn("p:grpSp"):
        raise PptxSkillsError(f'shape {shape.shape_id} "{shape.name}" is not a group')
    group = shape._element
    group_xfrm = _xfrm(group)
    if group_xfrm is None or group_xfrm.find(qn("a:chOff")) is None:
        raise PptxSkillsError(f'group "{shape.name}" has no child coordinate space')
    if int(group_xfrm.get("rot", "0")):
        raise PptxSkillsError(f'group "{shape.name}" is rotated; rotate it back before ungrouping')
    for child in [c for c in group if c.tag in _SHAPE_TAGS]:
        xfrm = _xfrm(child)
        if xfrm is not None:
            _convert(xfrm, group_xfrm)
        group.addprevious(child)
    group.getparent().remove(group)
    if timing.drop_shapes(slide._element, {shape.shape_id}):
        ctx.messages.append(f'slide {slide_number(ctx, slide)}: the animation of the group "{shape.name}" was removed with it')
    ctx.touched.add(slide.slide_id)
