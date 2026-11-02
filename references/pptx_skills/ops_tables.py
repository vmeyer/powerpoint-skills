"""Table ops: add a native table or fill one cloned from the examples, keeping its table style."""

import copy

from pptx.oxml.ns import qn
from pptx.util import Cm

from .changeset import Context, op, resolve_shape
from .colors import apply_color
from .errors import PptxSkillsError
from .ops_shapes import _check_name, _color, _snapped
from .ops_text import _shorthand, write_paragraphs
from .slides import resolve_slide


def _cell_paragraphs(text: str, font: dict) -> list[dict]:
    """The `text` shorthand for a cell: `**` adds bold, plain runs inherit the cell's and the table style's weight.

    Unlike a text box, a cell must not write `b=0` on plain runs: that would unbold header cells that the
    table style (or the cloned cell's first run) makes bold.
    """
    paragraphs = _shorthand(text, font)
    for paragraph in paragraphs:
        for run in paragraph["runs"]:
            if run.get("b") is False:
                del run["b"]
    return paragraphs


def _check_rectangular(rows: list[list[str]]) -> int:
    widths = {len(r) for r in rows}
    if len(widths) != 1:
        raise PptxSkillsError(f"rows must all have the same number of columns; found {sorted(widths)}")
    return widths.pop()


def _fill_cells(table, rows: list[list[str]], font: dict, allow_rgb: bool) -> None:
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            write_paragraphs(table.cell(r, c).text_frame, _cell_paragraphs(text, font), font, allow_rgb)


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
    box = _snapped(ctx, spec["box"])
    frame = slide.shapes.add_table(len(rows), cols, Cm(box.x), Cm(box.y), Cm(box.w), Cm(box.h))
    frame.name = spec["name"]
    table = frame.table
    table.first_row = spec.get("header", True)
    table.horz_banding = False  # banding is decoration the change set did not ask for
    if "style_id" in spec:  # without it python-pptx's built-in "Medium Style 2 - Accent 1" applies
        frame._element.graphic.graphicData.tbl.tblPr.find(qn("a:tableStyleId")).text = spec["style_id"]
    if widths is not None:
        for column, width in zip(table.columns, widths):
            column.width = Cm(width)
        frame.width = Cm(sum(widths))
    _fill_cells(table, rows, spec.get("font") or {}, allow_rgb)
    for (r, c), ref in fills.items():
        cell = table.cell(r, c)
        if ref is None:
            cell.fill.background()
        else:
            cell.fill.solid()
            apply_color(cell.fill.fore_color, ref)
    ctx.touched.add(slide.slide_id)


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
