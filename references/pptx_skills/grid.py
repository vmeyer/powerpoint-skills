"""`pptx grid`: aligned boxes inside the content area, so agents skip the grid arithmetic."""

from pathlib import Path

from .brief import load_brief
from .cli import emit
from .errors import PptxSkillsError
from .geometry import Box


def _spans(start: float, total: float, count: int, gutter: float) -> list[tuple[float, float]]:
    """(offset, size) per cell, rounded to 2 decimals; the last cell ends exactly at start + total."""
    size = (total - gutter * (count - 1)) / count
    if size <= 0:
        raise PptxSkillsError(f"gutter {gutter:g} cm leaves no room for {count} cells in {total:g} cm")
    edges = [(round(start + i * (size + gutter), 2), round(start + i * (size + gutter) + size, 2)) for i in range(count)]
    edges[-1] = (edges[-1][0], round(start + total, 2))
    return [(a, round(b - a, 2)) for a, b in edges]


def grid_boxes(area: Box, cols: int, rows: int = 1, gutter: float = 0.3, row_gutter: float | None = None) -> list[list[Box]]:
    row_gutter = gutter if row_gutter is None else row_gutter
    if cols < 1 or rows < 1:
        raise PptxSkillsError("cols and rows must be at least 1")
    if gutter < 0 or row_gutter < 0:
        raise PptxSkillsError("gutters must not be negative")
    xs = _spans(area.x, area.w, cols, gutter)
    ys = _spans(area.y, area.h, rows, row_gutter)
    return [[Box(x, y, w, h) for x, w in xs] for y, h in ys]


def _resolve_area(spec: str, deck_dir: Path) -> Box:
    brief = load_brief(deck_dir)
    if spec not in ("content", "full") and "," in spec:
        try:
            values = [float(v) for v in spec.split(",")]
        except ValueError:
            values = []
        if len(values) != 4:
            raise PptxSkillsError(f"--area must be content, full or x,y,w,h in cm; got {spec!r}")
        return Box(*values)
    if spec not in ("content", "full"):
        raise PptxSkillsError(f"--area must be content, full or x,y,w,h in cm; got {spec!r}")
    if brief is None:
        raise PptxSkillsError(f"no style brief in {deck_dir}; run `pptx init` / `pptx style` first")
    if spec == "full":
        return Box(0, 0, *brief["slide_size_cm"])
    if not brief["content_area"]:
        raise PptxSkillsError("the style brief has no content area; pass --area x,y,w,h")
    return Box(*brief["content_area"])


def _fmt(v: float) -> str:
    return f"{v:g}"


def _handle(args) -> int:
    area = _resolve_area(args.area, Path(args.deck_dir))
    rows = grid_boxes(area, args.cols, args.rows, args.gutter, args.row_gutter)
    lines = [" | ".join(f"{_fmt(b.x)},{_fmt(b.y)} {_fmt(b.w)}x{_fmt(b.h)}" for b in row) for row in rows]
    lines.append("Boxes are [x, y, w, h] in cm, ready for add_shape.")
    emit(args, {"area": area.as_list(), "rows": [[b.as_list() for b in row] for row in rows]}, "\n".join(lines))
    return 0


def register(subparsers) -> None:
    parser = subparsers.add_parser("grid", help="Compute aligned boxes inside the content area.")
    parser.add_argument("deck_dir", help="deck folder (with a style brief)")
    parser.add_argument("--cols", type=int, required=True)
    parser.add_argument("--rows", type=int, default=1)
    parser.add_argument("--gutter", type=float, default=0.3, help="column gutter in cm (default 0.3)")
    parser.add_argument("--row-gutter", type=float, default=None, help="row gutter in cm (default: --gutter)")
    parser.add_argument("--area", default="content", help="content (default), full, or x,y,w,h in cm")
    parser.add_argument("--json", action="store_true", help="machine-readable output")
    parser.set_defaults(handler=_handle)
