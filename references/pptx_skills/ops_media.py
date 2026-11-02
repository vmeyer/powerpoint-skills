"""Media ops: native charts colored from the theme, pictures cropped to fill, speaker notes."""

from pathlib import Path

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Cm, Pt

from .brief import DEFAULT_RULES
from .changeset import Context, op, resolve_shape
from .colors import apply_color, parse_color
from .errors import PptxSkillsError
from .ops_shapes import _check_name, _snapped
from .slides import resolve_slide

_CHART_TYPES = {
    "column_clustered": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column_stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "bar_clustered": XL_CHART_TYPE.BAR_CLUSTERED,
    "bar_stacked": XL_CHART_TYPE.BAR_STACKED,
    "line": XL_CHART_TYPE.LINE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
}
_LEGEND = {"bottom": XL_LEGEND_POSITION.BOTTOM, "right": XL_LEGEND_POSITION.RIGHT}
_ACCENTS = [MSO_THEME_COLOR[f"ACCENT_{i}"] for i in range(1, 7)]
_GRID = parse_color("tx1+85")  # thin, light value-axis gridlines in the text color
_GRID_PT = 0.5


def _min_size(ctx: Context) -> float:
    rules = (ctx.brief or {}).get("rules", {})
    return float(rules.get("min_size_pt", DEFAULT_RULES["min_size_pt"]))


def _accent(index: int, what: str) -> MSO_THEME_COLOR:
    if index >= len(_ACCENTS):
        raise PptxSkillsError(f"{what} needs more than the 6 theme accents (accent1 to accent6); reduce it to 6 or fewer")
    return _ACCENTS[index]


def _style_chart(chart, kind: str, spec: dict, size: float) -> None:
    chart.has_title = False
    chart.font.size = Pt(size)
    chart.font.color.theme_color = MSO_THEME_COLOR.TEXT_1  # family stays unset: the theme's minor font
    plot = chart.plots[0]
    series = list(plot.series)
    if kind == "doughnut":
        for i in range(len(spec["categories"])):
            point = series[0].points[i]
            point.format.fill.solid()
            point.format.fill.fore_color.theme_color = _accent(i, "a doughnut with this many categories")
    else:
        for i, ser in enumerate(series):
            color = _accent(i, "this many series")
            if kind == "line":
                ser.smooth = False
                ser.format.line.color.theme_color = color
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.theme_color = color
        axis = chart.value_axis
        axis.has_major_gridlines = True
        grid = axis.major_gridlines.format.line
        grid.width = Pt(_GRID_PT)
        apply_color(grid.color, _GRID)
    legend = spec.get("legend", "bottom" if kind == "doughnut" or len(series) > 1 else "none")
    chart.has_legend = legend != "none"
    if chart.has_legend:
        chart.legend.position = _LEGEND[legend]
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(size)
    if spec.get("labels"):
        plot.has_data_labels = True
        plot.data_labels.show_value = True
        plot.data_labels.font.size = Pt(size)


@op("add_chart")
def _add_chart(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    _check_name(slide, spec["name"])
    categories = spec["categories"]
    for ser in spec["series"]:
        if len(ser["values"]) != len(categories):
            raise PptxSkillsError(f'series "{ser["name"]}" has {len(ser["values"])} values for {len(categories)} categories; give one value per category')
    if spec["type"] == "doughnut" and len(spec["series"]) != 1:
        raise PptxSkillsError("a doughnut chart takes exactly one series")
    data = CategoryChartData()
    data.categories = categories
    for ser in spec["series"]:
        data.add_series(ser["name"], ser["values"])
    box = _snapped(ctx, spec["box"])
    frame = slide.shapes.add_chart(_CHART_TYPES[spec["type"]], Cm(box.x), Cm(box.y), Cm(box.w), Cm(box.h), data)
    frame.name = spec["name"]
    _style_chart(frame.chart, spec["type"], spec, _min_size(ctx))
    ctx.touched.add(slide.slide_id)


_SIGNATURES = {b"\x89PNG\r\n\x1a\n": "png", b"\xff\xd8\xff": "jpg"}


def _read_image(path_text: str) -> Path:
    path = Path(path_text)
    if not path.is_file():
        raise PptxSkillsError(f"picture not found: {path} (relative paths start at the current directory)")
    try:
        with path.open("rb") as handle:
            head = handle.read(8)
    except OSError as exc:
        raise PptxSkillsError(f"cannot read picture {path}: {exc}") from exc
    if not any(head.startswith(sig) for sig in _SIGNATURES):
        raise PptxSkillsError(f"{path.name} is not a PNG or JPG image (other formats such as SVG, GIF or WebP are not supported); convert it first")
    return path


def _crop_to_fill(pic, box_w: float, box_h: float) -> None:
    """Crop the dimension that overflows, evenly on both sides, so the image keeps its aspect ratio."""
    width, height = pic.image.size
    image_ratio, box_ratio = width / height, box_w / box_h
    pic.crop_left = pic.crop_right = pic.crop_top = pic.crop_bottom = 0
    if image_ratio > box_ratio:
        pic.crop_left = pic.crop_right = (1 - box_ratio / image_ratio) / 2
    elif image_ratio < box_ratio:
        pic.crop_top = pic.crop_bottom = (1 - image_ratio / box_ratio) / 2


@op("add_picture")
def _add_picture(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    path = _read_image(spec["path"])
    if "target" in spec:
        placeholder = resolve_shape(slide, spec["target"])
        if not hasattr(placeholder, "insert_picture"):
            raise PptxSkillsError(f'shape {placeholder.shape_id} "{placeholder.name}" is not a picture placeholder; use "box" to place a picture freely')
        if spec["name"] != placeholder.name:
            _check_name(slide, spec["name"])
        pic = placeholder.insert_picture(str(path))  # crops to fill the placeholder
    else:
        _check_name(slide, spec["name"])
        box = _snapped(ctx, spec["box"])
        pic = slide.shapes.add_picture(str(path), Cm(box.x), Cm(box.y), Cm(box.w), Cm(box.h))
        _crop_to_fill(pic, box.w, box.h)
    pic.name = spec["name"]
    pic._element.nvPicPr.cNvPr.set("descr", spec["alt"])  # python-pptx would put the file name here
    ctx.touched.add(slide.slide_id)


@op("set_notes")
def _set_notes(ctx: Context, spec: dict) -> None:
    slide = resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    slide.notes_slide.notes_text_frame.text = spec["text"]  # "\n" starts a new paragraph
    ctx.touched.add(slide.slide_id)
