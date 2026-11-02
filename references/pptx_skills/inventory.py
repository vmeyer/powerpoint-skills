"""Read a deck's slides and shapes into a plain dict: what is there now, where, in which role."""

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn

from . import cli
from .colors import read_color
from .errors import PptxSkillsError
from .geometry import Box, absolute_box
from .opc import open_presentation
from .roles import placeholder_role
from .units import EMU_PER_PT, emu_to_cm

_RUN_TAGS = (qn("a:r"), qn("a:br"), qn("a:fld"))
_AUTOFIT = {"noAutofit": "none", "spAutoFit": "shape", "normAutofit": "shrink"}
_TITLE_TYPES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.VERTICAL_TITLE)


def _number(value: float):
    return int(value) if float(value).is_integer() else round(value, 2)


def _kind(shape) -> str:
    tag = shape._element.tag.split("}")[1]
    if shape.is_placeholder:
        return "placeholder"
    if tag == "grpSp":
        return "group"
    if tag == "pic":
        return "picture"
    if tag == "cxnSp":
        return "connector"
    if tag == "graphicFrame":
        if getattr(shape, "has_table", False) and shape.has_table:
            return "table"
        if getattr(shape, "has_chart", False) and shape.has_chart:
            return "chart"
        return "other"
    if tag == "sp":
        c_nv = shape._element.find(qn("p:nvSpPr")).find(qn("p:cNvSpPr"))
        return "textbox" if c_nv is not None and c_nv.get("txBox") == "1" else "autoshape"
    return "other"


def _fill(sp_pr) -> str | None:
    if sp_pr is None:
        return None
    if sp_pr.find(qn("a:noFill")) is not None:
        return "none"
    solid = sp_pr.find(qn("a:solidFill"))
    if solid is not None:
        return read_color(solid)
    patt = sp_pr.find(qn("a:pattFill"))
    if patt is not None:
        parts = [patt.get("prst")]
        for name in ("a:fgClr", "a:bgClr"):
            el = patt.find(qn(name))
            parts.append(read_color(el) if el is not None else None)
        return "pattern:" + " ".join(str(p) for p in parts)
    if sp_pr.find(qn("a:gradFill")) is not None:
        return "gradient"
    if sp_pr.find(qn("a:blipFill")) is not None:
        return "image"
    return None


def _line(sp_pr):
    ln = sp_pr.find(qn("a:ln")) if sp_pr is not None else None
    if ln is None:
        return None
    if ln.find(qn("a:noFill")) is not None:
        return "none"
    solid = ln.find(qn("a:solidFill"))
    width = ln.get("w")
    return {
        "color": read_color(solid) if solid is not None else None,
        "width_pt": _number(int(width) / EMU_PER_PT) if width is not None else None,
    }


def _run(r) -> dict:
    if r.tag == qn("a:br"):
        return {"t": "\n", "br": True}
    run = {"t": r.findtext(qn("a:t")) or ""}
    if r.tag == qn("a:fld"):
        run["field"] = r.get("type")
    r_pr = r.find(qn("a:rPr"))
    if r_pr is not None:
        if r_pr.get("b") is not None:
            run["b"] = r_pr.get("b") in ("1", "true")
        if r_pr.get("i") is not None:
            run["i"] = r_pr.get("i") in ("1", "true")
        solid = r_pr.find(qn("a:solidFill"))
        if solid is not None and read_color(solid) is not None:
            run["color"] = read_color(solid)
        if r_pr.get("sz") is not None:
            run["size"] = _number(int(r_pr.get("sz")) / 100)
    return run


def _text(shape) -> dict | None:
    return _body_text(shape._element.find(qn("p:txBody")))


def _body_text(tx_body) -> dict | None:
    if tx_body is None:
        return None
    paragraphs = []
    for p in tx_body.findall(qn("a:p")):
        p_pr = p.find(qn("a:pPr"))
        level = int(p_pr.get("lvl", 0)) if p_pr is not None else 0
        paragraphs.append({"level": level, "runs": [_run(r) for r in p if r.tag in _RUN_TAGS]})
    body_pr = tx_body.find(qn("a:bodyPr"))
    autofit = None
    if body_pr is not None:
        autofit = next((v for tag, v in _AUTOFIT.items() if body_pr.find(qn(f"a:{tag}")) is not None), None)
    return {"paragraphs": paragraphs, "autofit": autofit}


def plain_text(text: dict | None) -> str:
    """A shape's inventory text as one string: runs joined, line breaks and paragraphs as "\\n"."""
    if not text:
        return ""
    return "\n".join("".join(r["t"] for r in p["runs"]) for p in text["paragraphs"])


def shape_text(shape) -> str:
    """The plain text of a python-pptx shape, exactly as the inventory reports it."""
    return plain_text(_text(shape))


def table_cell_texts(shape) -> list[list[str]]:
    """The plain text of every cell of a table shape, row by row (merged-away cells are empty)."""
    return [[plain_text(_body_text(tc.find(qn("a:txBody")))) for tc in tr.findall(qn("a:tc"))] for tr in shape._element.iter(qn("a:tr"))]


def _rotation(shape) -> float:
    xfrm = shape._element.xpath("./p:spPr/a:xfrm | ./p:grpSpPr/a:xfrm | ./p:xfrm")
    rot = int(xfrm[0].get("rot", 0)) / 60000 if xfrm else 0
    return _number(rot)


def _shape(shape, title_box: Box | None) -> dict:
    kind = _kind(shape)
    box = absolute_box(shape)
    el = shape._element
    sp_pr = el.find(qn("p:spPr"))
    prst = sp_pr.find(qn("a:prstGeom")) if sp_pr is not None else None
    role = "other"
    if kind == "placeholder":
        role = placeholder_role(shape.placeholder_format.type, box, title_box)
    elif kind in ("picture", "table", "chart"):
        role = kind
    c_nv_pr = el.xpath("./*[1]/p:cNvPr")[0]
    item = {
        "id": shape.shape_id,
        "name": shape.name,
        "kind": kind,
        "geometry": prst.get("prst") if prst is not None else None,
        "role": role,
        "box": box.as_list(),
        "rotation": _rotation(shape),
        "fill": _fill(sp_pr),
        "line": _line(sp_pr),
        "text": _text(shape),
        "alt": c_nv_pr.get("descr", ""),
        "children": [_shape(c, title_box) for c in shape.shapes] if kind == "group" else [],
    }
    if kind == "table":
        item["table"] = table_cell_texts(shape)
    return item


def _title_box(slide) -> Box | None:
    for shape in slide.placeholders:
        if shape.placeholder_format.type in _TITLE_TYPES:
            return absolute_box(shape)
    return None


def _role_text(shapes: list[dict], role: str) -> str | None:
    for shape in shapes:
        if shape["kind"] == "placeholder" and shape["role"] == role:
            return plain_text(shape["text"])
    return None


def inventory(prs, slides: list[int] | None = None) -> dict:
    count = len(prs.slides)
    wanted = slides if slides is not None else list(range(1, count + 1))
    bad = [n for n in wanted if not 1 <= n <= count]
    if bad:
        raise PptxSkillsError(f"slide {bad[0]} does not exist; the deck has {count} slides")
    result = []
    for number in wanted:
        slide = prs.slides[number - 1]
        title_box = _title_box(slide)
        shapes = [_shape(s, title_box) for s in slide.shapes]
        frame = slide.notes_slide.notes_text_frame if slide.has_notes_slide else None  # None: no body placeholder
        notes = frame.text if frame is not None else ""
        result.append(
            {
                "number": number,
                "id": slide.slide_id,
                "layout": slide.slide_layout.name,
                "hidden": slide._element.get("show") in ("0", "false"),
                "title": _role_text(shapes, "title"),
                "kicker": _role_text(shapes, "kicker"),
                "notes": notes,
                "shapes": shapes,
            }
        )
    return {
        "slide_size_cm": [emu_to_cm(prs.slide_width), emu_to_cm(prs.slide_height)],
        "slides": result,
    }


def _g(value: float) -> str:
    return f"{value:g}"


def _lines(shapes: list[dict], depth: int) -> list[str]:
    out = []
    for s in shapes:
        x, y, w, h = s["box"]
        line = f"{'  ' * depth}- [{s['id']}] {s['name']} {s['kind']}/{s['role']} @ {_g(x)},{_g(y)} {_g(w)}×{_g(h)}"
        text = " / ".join(t for t in plain_text(s["text"]).split("\n") if t)
        if text:
            line += f': "{text if len(text) <= 60 else text[:59] + "…"}"'
        out.append(line)
        out.extend(_lines(s["children"], depth + 1))
    return out


def to_markdown(inv: dict) -> str:
    out = []
    for slide in inv["slides"]:
        hidden = " (hidden)" if slide.get("hidden") else ""
        out.append(f"## Slide {slide['number']} (id {slide['id']}) — {slide['layout']}{hidden}")
        out.extend(_lines(slide["shapes"], 0))
        out.append("")
    return "\n".join(out).rstrip()


def outline(inv: dict) -> dict:
    """Number, stable id, kicker and title per slide: what a storyline needs from an existing deck."""
    return {"slides": [{key: slide[key] for key in ("number", "id", "kicker", "title")} for slide in inv["slides"]]}


def to_outline_text(data: dict) -> str:
    """One line per slide, `N (id I) | kicker | title`; "—" where there is no text."""

    def one_line(text: str | None) -> str:
        return " ".join((text or "").split()) or "—"

    return "\n".join(f"{s['number']} (id {s['id']}) | {one_line(s['kicker'])} | {one_line(s['title'])}" for s in data["slides"])


def _handle(args) -> int:
    slides = None
    if args.slides:
        try:
            slides = [int(n) for n in args.slides.split(",")]
        except ValueError:
            raise PptxSkillsError(f"--slides expects numbers like 1,3, got {args.slides!r}") from None
    inv = inventory(open_presentation(args.deck), slides)
    if args.outline:
        data = outline(inv)
        cli.emit(args, data, to_outline_text(data))
    else:
        cli.emit(args, inv, to_markdown(inv))
    return cli.OK


def register(subparsers) -> None:
    parser = subparsers.add_parser("inventory", help="List the slides and shapes of a deck.")
    parser.add_argument("deck", help="path to a .pptx or .potx file")
    parser.add_argument("--slides", help="comma-separated slide numbers, e.g. 1,3")
    parser.add_argument("--outline", action="store_true", help="one line per slide: number, id, kicker, title")
    parser.add_argument("--json", action="store_true", help="machine-readable output")
    parser.set_defaults(handler=_handle)
