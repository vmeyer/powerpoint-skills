"""Table ops: add a native table or fill one cloned from the examples, keeping its table style."""

import copy
import re

from pptx.oxml.ns import qn
from pptx.util import Cm

from .brief import default_table_style, table_style_catalog
from .changeset import Context, op, resolve_shape
from .colors import apply_color
from .errors import PptxSkillsError
from .ops_shapes import _check_name, _color, _snapped
from .ops_text import _shorthand, write_paragraphs
from .slides import resolve_slide


_GUID = re.compile(r"^\{[0-9A-Fa-f]{8}(-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}$")


def _style_id(prs, style: str | None) -> str | None:
    """The style id for `add_table`: a GUID as given, a name of the deck's tableStyles part, else the deck's default."""
    if style is None:
        return default_table_style(prs)
    if _GUID.match(style):
        return style
    catalog = table_style_catalog(prs)
    for entry in catalog:
        if (entry["name"] or "").casefold() == style.casefold():
            return entry["id"]
    names = ", ".join(f'"{e["name"]}"' for e in catalog) or "none defined"
    raise PptxSkillsError(f'unknown table style "{style}"; styles of this deck: {names} (or give a style GUID)')


def _check_rectangular(rows: list[list[str]]) -> int:
    widths = {len(r) for r in rows}
    if len(widths) != 1:
        raise PptxSkillsError(f"rows must all have the same number of columns; found {sorted(widths)}")
    return widths.pop()


def _fill_cells(table, rows: list[list[str]], font: dict, allow_rgb: bool) -> None:
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            # the shorthand rule of set_text: without `**` the cell takes its weight from the table style
            write_paragraphs(table.cell(r, c).text_frame, _shorthand(text, font), font, allow_rgb)


@op("add_table")
def _add_table(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    _check_name(slide, spec["name"])
    rows = spec["rows"]
    cols = _check_rectangular(rows)
    widths = spec.get("col_widths")
    if widths is not None and len(widths) != cols:
        raise PptxSkillsError(f"col_widths lists {len(widths)} widths for {cols} columns")
    allow_rgb = bool(spec.get("allow_rgb"))
    fills = {}
    for key, color in spec.get("fills", {}).items():
        r, c = (int(v) for v in key.split(","))
        if r >= len(rows) or c >= cols:
            raise PptxSkillsError(f'fills "{key}" is outside the {len(rows)}x{cols} table (row,column are 0-based)')
        fills[(r, c)] = _color(color, allow_rgb, f'fills "{key}"')
    style_id = _style_id(ctx.prs, spec.get("style"))
    box = _snapped(ctx, spec["box"])
    frame = slide.shapes.add_table(len(rows), cols, Cm(box.x), Cm(box.y), Cm(box.w), Cm(box.h))
    frame.name = spec["name"]
    table = frame.table
    table.first_row = spec.get("header", True)
    table.horz_banding = False  # banding is decoration the change set did not ask for
    if style_id:  # otherwise python-pptx's built-in "Medium Style 2 - Accent 1" stays
        frame._element.graphic.graphicData.tbl.tblPr.find(qn("a:tableStyleId")).text = style_id
    if widths is not None:
        for column, width in zip(table.columns, widths):
            column.width = Cm(width)
        frame.width = sum(column.width for column in table.columns)
    _fill_cells(table, rows, spec.get("font") or {}, allow_rgb)
    for (r, c), ref in fills.items():
        cell = table.cell(r, c)
        if ref is None:
            cell.fill.background()
        else:
            cell.fill.solid()
            apply_color(cell.fill.fore_color, ref)
    ctx.touched.add(slide.slide_id)


def _check_spans(name: str, trs: list, wanted: int) -> None:
    """Refuse a resize that would copy or cut a vertical merge."""
    def span(tc) -> int:
        return int(tc.get("rowSpan", 1))

    if wanted > len(trs) and any(tc.get("vMerge") or span(tc) > 1 for tc in trs[-1].findall(qn("a:tc"))):
        raise PptxSkillsError(f'the last row of table "{name}" is part of a vertically merged cell; unmerge it in PowerPoint before adding rows')
    for i, tr in enumerate(trs[:wanted]):
        if any(i + span(tc) > wanted for tc in tr.findall(qn("a:tc"))):
            raise PptxSkillsError(f'removing rows would cut a vertically merged cell of table "{name}"; unmerge it in PowerPoint or keep at least {max(i + span(tc) for tc in tr.findall(qn("a:tc")))} rows')


@op("set_table")
def _set_table(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    frame = resolve_shape(slide, spec["target"])
    if not getattr(frame, "has_table", False) or not frame.has_table:
        raise PptxSkillsError(f'shape {frame.shape_id} "{frame.name}" is not a table')
    rows = spec["rows"]
    cols = _check_rectangular(rows)
    table = frame.table
    if cols != len(table.columns):
        raise PptxSkillsError(f'the rows have {cols} columns but table "{frame.name}" has {len(table.columns)}; set_table does not add or remove columns')
    tbl = frame._element.graphic.graphicData.tbl
    trs = tbl.tr_lst
    _check_spans(frame.name, trs, len(rows))
    while len(trs) > len(rows):
        tbl.remove(trs.pop())
    while len(trs) < len(rows):
        clone = copy.deepcopy(trs[-1])
        for ext in clone.findall(qn("a:extLst")):  # a copied row id would collide with the original's
            clone.remove(ext)
        trs[-1].addnext(clone)
        trs.append(clone)
    frame.height = sum(int(tr.get("h")) for tr in trs)
    _fill_cells(table, rows, spec.get("font") or {}, bool(spec.get("allow_rgb")))
    ctx.touched.add(slide.slide_id)
