"""The style brief: facts about a master (theme, text styles, grid, layouts, example usage, check rules)."""

import json
import statistics
from collections import Counter
from pathlib import Path

from lxml import etree
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

from .colors import read_color
from .geometry import Box, absolute_box
from .inventory import inventory
from .roles import placeholder_role
from .units import emu_to_cm, guide_to_cm

DEFAULT_RULES: dict = {
    "min_size_pt": 11,
    "min_size_small_pt": 9,
    "max_lines": 3,
    "max_title_lines": 2,
    "max_bullets": 5,
    "near_align_cm": 0.3,
    "snap_cm": 0.05,
    "title_min_words": 4,
}

BLOCK_START = "<!-- agent:building-blocks -->"
BLOCK_END = "<!-- /agent:building-blocks -->"
_BLOCK_HINT = "Reusable building blocks found in the example slides (the agent keeps this section)."

_P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main"
_THEME_COLORS = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink")
_TITLE_TYPES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.VERTICAL_TITLE)
_BODY_ROLES = ("body",)
_IGNORED_ROLES = ("footer", "number", "date")
_MAX_GAP_CM = 3.0
_MIN_SPAN_OVERLAP = 0.5


def _number(value: float):
    return int(value) if float(value).is_integer() else round(value, 2)


def _theme(prs) -> dict:
    root = etree.fromstring(prs.slide_master.part.part_related_by(RT.THEME).blob)
    scheme = root.find(".//" + qn("a:clrScheme"))
    colors = {}
    for role in _THEME_COLORS:
        el = scheme.find(qn(f"a:{role}")) if scheme is not None else None
        if el is not None and read_color(el) is not None:
            colors[role] = read_color(el)
    fonts = root.find(".//" + qn("a:fontScheme"))
    out = {}
    for key, tag in (("major", "a:majorFont"), ("minor", "a:minorFont")):
        latin = fonts.find(f"{qn(tag)}/{qn('a:latin')}") if fonts is not None else None
        out[key] = latin.get("typeface") if latin is not None else None
    return {"colors": colors, "fonts": out}


def _text_styles(prs) -> dict:
    tx_styles = prs.slide_master._element.find(qn("p:txStyles"))

    def sizes(style_tag: str) -> list:
        style = tx_styles.find(qn(style_tag)) if tx_styles is not None else None
        out = []
        for level in range(1, 10):
            def_rpr = style.find(f"{qn(f'a:lvl{level}pPr')}/{qn('a:defRPr')}") if style is not None else None
            if def_rpr is not None and def_rpr.get("sz"):
                out.append(_number(int(def_rpr.get("sz")) / 100))
        return out

    title = sizes("p:titleStyle")
    return {"title_pt": title[0] if title else None, "body_pt": sizes("p:bodyStyle")}


def _guides(prs) -> dict:
    vertical, horizontal = [], []
    for root in (prs.slide_master._element, prs.part._element):
        for guide in root.iter(f"{{{_P15}}}guide"):
            pos = guide.get("pos")
            if pos is None:
                continue
            target = horizontal if guide.get("orient") == "horz" else vertical
            target.append(round(guide_to_cm(int(pos)), 2))
    return {"vertical": sorted(vertical), "horizontal": sorted(horizontal)}


def _layout_title_box(layout) -> Box | None:
    for ph in layout.placeholders:
        if ph.placeholder_format.type in _TITLE_TYPES:
            return absolute_box(ph)
    return None


def _layouts(prs) -> list[dict]:
    result = []
    for layout in prs.slide_layouts:
        title_box = _layout_title_box(layout)
        placeholders = []
        for ph in layout.placeholders:
            box = absolute_box(ph)
            fmt = ph.placeholder_format
            placeholders.append(
                {
                    "idx": fmt.idx,
                    "type": fmt.type.name.lower() if fmt.type is not None else "other",
                    "role": placeholder_role(fmt.type, box, title_box),
                    "box": box.as_list(),
                    "prompt": ph.text_frame.text if ph.has_text_frame else "",
                }
            )
        result.append({"name": layout.name, "placeholders": placeholders})
    return result


def _area_from_layouts(layouts: list[dict]) -> list[float] | None:
    best, best_area = None, -1.0
    for layout in layouts:
        roles = [p["role"] for p in layout["placeholders"] if p["role"] not in _IGNORED_ROLES]
        if sorted(roles) != ["body", "title"]:
            continue
        box = next(p["box"] for p in layout["placeholders"] if p["role"] == "body")
        if box[2] * box[3] > best_area:
            best, best_area = box, box[2] * box[3]
    return best


def _content_area(prs, guides: dict, layouts: list[dict]) -> tuple[list[float] | None, str]:
    if len(guides["vertical"]) == 2 and len(guides["horizontal"]) == 2:
        (left, right), (top, bottom) = guides["vertical"], guides["horizontal"]
        return [left, top, round(right - left, 2), round(bottom - top, 2)], "guides"
    return _area_from_layouts(layouts), "layouts"


def _is_color(fill) -> bool:
    return isinstance(fill, str) and fill != "none" and not fill.startswith(("pattern:", "gradient", "image"))


def _walk(shapes: list[dict]):
    for shape in shapes:
        yield shape
        yield from _walk(shape["children"])


def _gaps(shapes: list[dict]) -> tuple[list[float], list[float]]:
    horizontal, vertical = [], []
    boxes = [Box(*s["box"]) for s in shapes]
    for a in boxes:
        for b in boxes:
            if a is b:
                continue
            gap = b.x - a.right
            overlap = min(a.bottom, b.bottom) - max(a.y, b.y)
            if 0 < gap < _MAX_GAP_CM and overlap >= _MIN_SPAN_OVERLAP * min(a.h, b.h) > 0:
                horizontal.append(gap)
            gap = b.y - a.bottom
            overlap = min(a.right, b.right) - max(a.x, b.x)
            if 0 < gap < _MAX_GAP_CM and overlap >= _MIN_SPAN_OVERLAP * min(a.w, b.w) > 0:
                vertical.append(gap)
    return horizontal, vertical


def _sibling_lists(shapes: list[dict]):
    yield shapes
    for shape in shapes:
        if shape["children"]:
            yield from _sibling_lists(shape["children"])


def _corner_radii(slide) -> set[float]:
    radii = set()

    def visit(shapes):
        for shape in shapes:
            geom = shape._element.find(f"{qn('p:spPr')}/{qn('a:prstGeom')}")
            if geom is not None and geom.get("prst") == "roundRect":
                adj = 16667
                for gd in geom.iterfind(f"{qn('a:avLst')}/{qn('a:gd')}"):
                    if gd.get("name") == "adj" and gd.get("fmla", "").startswith("val "):
                        adj = int(gd.get("fmla")[4:])
                box = absolute_box(shape)
                radii.add(round(adj / 100000 * min(box.w, box.h), 2))
            if shape._element.tag == qn("p:grpSp"):
                visit(shape.shapes)

    visit(slide.shapes)
    return radii


def _examples(examples) -> dict:
    inv = inventory(examples)
    entries, fills_total, text_total = [], Counter(), Counter()
    gaps_h, gaps_v, radii = [], [], set()
    for slide, source in zip(inv["slides"], examples.slides):
        fills, text_colors, sizes = Counter(), Counter(), set()
        shapes = list(_walk(slide["shapes"]))
        for shape in shapes:
            if _is_color(shape["fill"]):
                fills[shape["fill"]] += 1
            for paragraph in (shape["text"] or {"paragraphs": []})["paragraphs"]:
                for run in paragraph["runs"]:
                    if run.get("color"):
                        text_colors[run["color"]] += 1
                    if "size" in run:
                        sizes.add(run["size"])
        for siblings in _sibling_lists(slide["shapes"]):
            h, v = _gaps(siblings)
            gaps_h += h
            gaps_v += v
        radii |= _corner_radii(source)
        fills_total.update(fills)
        text_total.update(text_colors)
        entries.append(
            {
                "number": slide["number"],
                "layout": slide["layout"],
                "shape_count": len(shapes),
                "fills": dict(fills),
                "text_colors": dict(text_colors),
                "sizes_pt": sorted(sizes),
                "table_styles": sorted({el.text for el in source._element.iter(qn("a:tableStyleId")) if el.text}),
            }
        )
    median = lambda values: round(statistics.median(values), 2) if values else None  # noqa: E731
    return {
        "examples": entries,
        "color_usage": {"fills": dict(fills_total), "text": dict(text_total)},
        "typical_gap_cm": {"horizontal": median(gaps_h), "vertical": median(gaps_v)},
        "corner_radii": sorted(radii),
    }


def build_brief(master, examples=None, rules_override: dict | None = None) -> dict:
    guides = _guides(master)
    layouts = _layouts(master)
    area, source = _content_area(master, guides, layouts)
    usage = _examples(examples) if examples is not None else {
        "examples": [],
        "color_usage": {"fills": {}, "text": {}},
        "typical_gap_cm": {"horizontal": None, "vertical": None},
        "corner_radii": [],
    }
    return {
        "slide_size_cm": [emu_to_cm(master.slide_width), emu_to_cm(master.slide_height)],
        "theme": _theme(master),
        "text_styles": _text_styles(master),
        "guides": guides,
        "content_area": area,
        "content_area_source": source,
        "layouts": layouts,
        **usage,
        "rules": {**DEFAULT_RULES, **(rules_override or {})},
    }


def _existing_block(md_path: Path) -> str:
    if md_path.is_file():
        text = md_path.read_text(encoding="utf-8")
        start, end = text.find(BLOCK_START), text.find(BLOCK_END)
        if start != -1 and end > start:
            return text[start + len(BLOCK_START) : end]
    return f"\n{_BLOCK_HINT}\n"


def _g(value) -> str:
    return "-" if value is None else f"{value:g}"


def _markdown(brief: dict, block: str) -> str:
    theme = brief["theme"]
    w, h = brief["slide_size_cm"]
    area = brief["content_area"]
    lines = [
        "# Style brief",
        "",
        f"Slide size: {_g(w)} x {_g(h)} cm",
        f"Fonts: {theme['fonts']['major']} (headings), {theme['fonts']['minor']} (body)",
        "",
        "## Theme colors",
        *[f"- {role}: {value}" for role, value in theme["colors"].items()],
        "",
        "## Text styles",
        f"- Title: {_g(brief['text_styles']['title_pt'])} pt",
        f"- Body levels: {', '.join(_g(s) for s in brief['text_styles']['body_pt'])} pt",
        "",
        "## Content area",
        f"- x, y, w, h: {', '.join(_g(v) for v in area) if area else 'unknown'} cm (from {brief['content_area_source']})",
        f"- Typical gap: horizontal {_g(brief['typical_gap_cm']['horizontal'])} cm, vertical {_g(brief['typical_gap_cm']['vertical'])} cm",
        f"- Corner radii: {', '.join(_g(r) for r in brief['corner_radii']) or '-'} cm",
        "",
        "## Layouts",
    ]
    for layout in brief["layouts"]:
        roles = ", ".join(p["role"] for p in layout["placeholders"] if p["role"] not in _IGNORED_ROLES) or "no placeholders"
        lines.append(f"- {layout['name']}: {roles}")
    lines += ["", "## Color usage in the examples"]
    for kind in ("fills", "text"):
        usage = brief["color_usage"][kind]
        lines.append(f"- {kind}: " + (", ".join(f"{ref} x{n}" for ref, n in usage.items()) or "-"))
    lines += ["", "## Examples"]
    for ex in brief["examples"]:
        lines.append(f"- Slide {ex['number']} ({ex['layout']}): {ex['shape_count']} shapes, sizes {', '.join(_g(s) for s in ex['sizes_pt']) or '-'} pt")
    lines += ["", "## Check rules", *[f"- {key}: {value}" for key, value in brief["rules"].items()]]
    lines += ["", "## Building blocks", BLOCK_START + block + BLOCK_END, ""]
    return "\n".join(lines)


def write_brief(brief: dict, style_dir: Path) -> None:
    style_dir = Path(style_dir)
    style_dir.mkdir(parents=True, exist_ok=True)
    block = _existing_block(style_dir / "brief.md")
    (style_dir / "brief.json").write_text(json.dumps(brief, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    (style_dir / "brief.md").write_text(_markdown(brief, block), encoding="utf-8")


def load_brief(deck_dir: Path) -> dict | None:
    path = Path(deck_dir) / "style" / "brief.json"
    return json.loads(path.read_text(encoding="utf-8")) if path.is_file() else None
