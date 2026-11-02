"""Geometry rules: shapes stay on the slide, inside the content area, apart and aligned."""

from ..geometry import Box
from ..units import EMU_PER_CM
from . import rule

OFF_SLIDE_TOLERANCE_CM = 0.01
OVERLAP_MIN_CM2 = 0.01
ALIGN_NOISE_CM = 0.02  # guides are stored in 1/8 pt, so edges can sit ~0.005 cm off (R20)
GAP_MAX_CM = 3.0
GAP_SPREAD_CM = 0.1
ROW_SPAN_OVERLAP = 0.5
_SMALL_ROLES = ("footer", "number", "date")
_CHROME_ROLES = ("title", "kicker") + _SMALL_ROLES
_LINES = ("connector",)


def _flat(items):
    for item in items:
        yield item
        yield from _flat(item["children"])


def _levels(items):
    """Every list of siblings: the top level and the children of each group."""
    yield items
    for item in items:
        if item["children"]:
            yield from _levels(item["children"])


def _plain(item) -> bool:
    return item["kind"] not in _LINES and item["role"] not in _SMALL_ROLES and item["box"][2] > 0 and item["box"][3] > 0


def _has_text(item) -> bool:
    paragraphs = (item["text"] or {}).get("paragraphs", [])
    return any(run["t"].strip() for p in paragraphs for run in p["runs"])


@rule("off-slide", "geometry", "error")
def off_slide(ctx, slide):
    width, height = ctx.prs.slide_width / EMU_PER_CM, ctx.prs.slide_height / EMU_PER_CM
    tol = OFF_SLIDE_TOLERANCE_CM
    out = []
    for item in _flat(slide.info["shapes"]):
        if item["kind"] == "group":
            continue
        x, y, w, h = item["box"]
        if x < -tol or y < -tol or x + w > width + tol or y + h > height + tol:
            out.append(ctx.finding(slide, item, f"the shape reaches beyond the slide ({width:.2f} × {height:.2f} cm) at {x:g}, {y:g}, {w:g} × {h:g} cm", "move or resize the shape so it lies fully on the slide"))
    return out


@rule("outside-content", "geometry", "warning")
def outside_content(ctx, slide):
    if not ctx.content_area:
        return []
    area = Box(*ctx.content_area)
    out = []
    for item in _flat(slide.info["shapes"]):
        box = Box(*item["box"])
        if item["kind"] == "group" or item["role"] in _CHROME_ROLES or box.contains(area, 0.01):
            continue
        if not area.contains(box, 0.01):
            out.append(ctx.finding(slide, item, f"the shape leaves the content area (x {area.x:g}–{area.right:g}, y {area.y:g}–{area.bottom:g} cm)", "move or resize the shape into the content area"))
    return out


@rule("overlap", "geometry", "warning")
def overlap(ctx, slide):
    shapes = [i for i in _flat(slide.info["shapes"]) if _plain(i) and i["kind"] != "group" and _has_text(i)]
    out = []
    for n, first in enumerate(shapes):
        a = Box(*first["box"])
        for second in shapes[n + 1 :]:
            b = Box(*second["box"])
            area = a.intersects(b)
            if area > OVERLAP_MIN_CM2 and not a.contains(b) and not b.contains(a):
                out.append(ctx.finding(slide, second, f"“{first['name']}” and “{second['name']}” overlap by {area:.2f} cm²", "move one of the two shapes or shrink them so the texts do not collide"))
    return out


@rule("near-aligned", "geometry", "warning")
def near_aligned(ctx, slide):
    limit = ctx.rules["near_align_cm"]
    out = []
    for level in _levels(slide.info["shapes"]):
        shapes = [i for i in level if _plain(i)]
        for n, first in enumerate(shapes):
            a = Box(*first["box"])
            for second in shapes[n + 1 :]:
                b = Box(*second["box"])
                edges = [
                    (name, abs(getattr(a, attr) - getattr(b, attr)))
                    for name, attr in (("left", "x"), ("right", "right"), ("top", "y"), ("bottom", "bottom"))
                ]
                near = [(name, d) for name, d in edges if ALIGN_NOISE_CM - 1e-9 <= d <= limit + 1e-9]
                if near:
                    name, d = near[0]
                    out.append(ctx.finding(slide, second, f"the {name} edges of “{first['name']}” and “{second['name']}” differ by {d:.2f} cm", f"align the {name} edge of “{second['name']}” with “{first['name']}” or move it clearly apart"))
    return out


def _rows(shapes):
    """Groups of at least three shapes whose vertical extents overlap by half, each row once."""
    seen: list[frozenset] = []
    for seed in shapes:
        s = Box(*seed["box"])
        row = []
        for item in shapes:
            b = Box(*item["box"])
            shared = min(s.bottom, b.bottom) - max(s.y, b.y)
            if shared >= ROW_SPAN_OVERLAP * min(s.h, b.h):
                row.append(item)
        key = frozenset(i["id"] for i in row)
        if len(row) >= 3 and not any(key <= other for other in seen):
            seen.append(key)
            yield sorted(row, key=lambda i: i["box"][0])


@rule("uneven-gaps", "geometry", "warning")
def uneven_gaps(ctx, slide):
    out = []
    for level in _levels(slide.info["shapes"]):
        for row in _rows([i for i in level if _plain(i)]):
            gaps = [round(b["box"][0] - (a["box"][0] + a["box"][2]), 2) for a, b in zip(row, row[1:])]
            if min(gaps) < 0 or max(gaps) >= GAP_MAX_CM or max(gaps) - min(gaps) <= GAP_SPREAD_CM + 1e-9:
                continue
            names = ", ".join(f"“{i['name']}”" for i in row)
            listing = ", ".join(f"{g:g}" for g in gaps)
            out.append(ctx.finding(slide, row[0], f"the gaps in the row {names} are uneven ({listing} cm)", "space the shapes evenly, for example with `pptx distribute`"))
    return out
