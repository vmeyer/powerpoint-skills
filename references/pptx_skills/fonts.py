"""Effective text styles, installed-font lookup and wrapped-text measurement.

Most runs carry no size or font of their own; PowerPoint resolves each run attribute along an
inheritance chain where the first explicit value wins:

    run rPr → paragraph pPr/defRPr → shape lstStyle → layout placeholder lstStyle
    → master placeholder lstStyle → master txStyles (titleStyle for title/ctrTitle placeholders,
    bodyStyle for other placeholders, otherStyle for non-placeholders)
    → presentation defaultTextStyle → 18 pt

At every list-style step the paragraph level (`a:pPr lvl`, default 0) selects `a:lvl{lvl+1}pPr`.
Paragraph spacing (`lnSpc`, `spcBef`, `spcAft`) and the left margin (`marL`) follow the same chain
from `a:pPr` on. Body properties (insets, wrapping, autofit) follow the shape's, the layout
placeholder's and the master placeholder's `a:bodyPr`. Malformed numbers count as absent, so the
next source in the chain applies.
"""

import math
import os
import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from PIL import ImageFont
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

from .geometry import absolute_box
from .units import EMU_PER_CM, EMU_PER_PT

DEFAULT_SIZE_PT = 18.0
FALLBACK_LINE_HEIGHT = 1.2  # single line height as a multiple of the size when the font is missing
_SCALE = 10  # fonts are measured at 10x size for sub-point precision

# ECMA-376 defaults: 0.1 in left/right, 0.05 in top/bottom (0.25 cm / 0.13 cm).
_DEFAULT_INSETS = {"lIns": 91440, "tIns": 45720, "rIns": 91440, "bIns": 45720}
_AUTOFIT = {qn("a:noAutofit"): "none", qn("a:normAutofit"): "shrink", qn("a:spAutoFit"): "shape"}

_OFFICE_FONTS = "Microsoft PowerPoint.app/Contents/Resources/DFonts"
_FONT_DIRS = (
    "/System/Library/Fonts",
    "/System/Library/Fonts/Supplemental",
    "/Library/Fonts",
    "~/Library/Fonts",
    "/usr/share/fonts",
    "~/.fonts",
    f"/Applications/{_OFFICE_FONTS}",  # fonts that ship only with Office, such as Calibri
    f"~/Applications/{_OFFICE_FONTS}",
)
_FONT_SUFFIXES = frozenset({".ttf", ".otf", ".ttc", ".otc"})
_COLLECTION_SUFFIXES = frozenset({".ttc", ".otc"})
_PLAIN_STYLE_WORDS = frozenset({"regular", "roman", "normal", "book", "plain"})
_RIBBI_WORDS = _PLAIN_STYLE_WORDS | {"bold", "italic", "oblique"}

_TITLE_TYPES = frozenset({"title", "ctrTitle"})
# Layout placeholders inherit from the master placeholder of the corresponding type; the master
# only carries title, body, date, footer, slide number (and header) placeholders.
_MASTER_TYPES = {"title": "title", "ctrTitle": "title", "dt": "dt", "ftr": "ftr", "sldNum": "sldNum", "hdr": "hdr"}

_RUN_TAGS = (qn("a:r"), qn("a:fld"), qn("a:br"))
# Hard breaks inside run text: vertical tab (python-pptx's line-break marker, stored escaped) and newline.
_BREAK = re.compile(r"_x000[bB]_|[\v\n]")
_TOKEN = re.compile(r"\S+|\s+")


@dataclass(frozen=True)
class RunStyle:
    family: str
    size_pt: float
    bold: bool
    italic: bool


@dataclass
class Measurement:
    """Wrapped text of one shape, measured at the declared sizes.

    `height_cm` includes the top and bottom insets, so it compares directly with the shape's
    height. `autofit` is the resolved mode: "none", "shrink" (normAutofit) or "shape" (spAutoFit);
    `font_scale` and `line_spacing_reduction` are normAutofit's stored values (1.0 and 0.0 otherwise).
    """

    lines_per_paragraph: list[int]
    height_cm: float
    widest_word_cm: float
    fonts_missing: set[str]
    autofit: str = "none"
    font_scale: float = 1.0
    line_spacing_reduction: float = 0.0


def _int(value: str | None) -> int | None:
    """An integer attribute; None when absent or malformed."""
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _percent(value: str | None) -> float | None:
    """A percentage as a fraction: 1/1000 % as an integer ("62500") or a string ("62.5%")."""
    if value is None:
        return None
    try:
        return float(value[:-1]) / 100 if value.endswith("%") else int(value) / 100000
    except ValueError:
        return None


# --- inheritance -------------------------------------------------------------------------------


def _ph_type(ph) -> str:
    return ph.get("type", "obj")


def _placeholder_elements(owner) -> list:
    return [shape._element for shape in owner.placeholders]


def _layout_placeholder(ph, layout):
    """The layout placeholder a slide placeholder inherits from: same idx first, else same type."""
    elements = _placeholder_elements(layout)
    idx = _int(ph.get("idx"))
    if idx is not None:
        found = next((el for el in elements if (_int(el.ph.get("idx")) or 0) == idx), None)
        if found is not None:
            return found
    return next((el for el in elements if _ph_type(el.ph) == _ph_type(ph)), None)


def _master_placeholder(ph_type: str, master):
    wanted = _MASTER_TYPES.get(ph_type, "body")
    return next((el for el in _placeholder_elements(master) if _MASTER_TYPES.get(_ph_type(el.ph), "body") == wanted), None)


def _tx_child(element, tag):
    tx_body = element.find(qn("p:txBody")) if element is not None else None
    return tx_body.find(qn(tag)) if tx_body is not None else None


@lru_cache(maxsize=16)
def _theme_fonts_from(blob: bytes) -> tuple[str | None, str | None]:
    root = parse_xml(blob)
    fonts = []
    for tag in ("a:majorFont", "a:minorFont"):
        latin = root.find(f".//{qn(tag)}/{qn('a:latin')}")
        fonts.append(latin.get("typeface") if latin is not None else None)
    return fonts[0], fonts[1]


def _theme_fonts(master) -> tuple[str | None, str | None]:
    try:
        theme = master.part.part_related_by(RT.THEME)
    except KeyError:
        return None, None
    return _theme_fonts_from(theme.blob)


@dataclass(frozen=True)
class _Body:
    insets: dict[str, int]
    wrap: bool
    autofit: str
    font_scale: float
    line_spacing_reduction: float


class _Inheritance:
    """Everything a shape's text inherits from, collected once per shape."""

    def __init__(self, shape, slide):
        element = shape._element
        layout = slide.slide_layout
        master = layout.slide_master
        ph = element.ph
        layout_ph = master_ph = ph_type = None
        if ph is not None:
            layout_ph = _layout_placeholder(ph, layout)
            ph_type = ph.get("type") or (layout_ph.ph.get("type") if layout_ph is not None else None) or "obj"
            master_ph = _master_placeholder(ph_type, master)
            style_tag = "p:titleStyle" if ph_type in _TITLE_TYPES else "p:bodyStyle"
        else:
            style_tag = "p:otherStyle"
        tx_styles = master._element.find(qn("p:txStyles"))
        presentation = slide.part.package.presentation_part._element
        list_styles = [_tx_child(el, "a:lstStyle") for el in (element, layout_ph, master_ph)]
        list_styles.append(tx_styles.find(qn(style_tag)) if tx_styles is not None else None)
        list_styles.append(presentation.find(qn("p:defaultTextStyle")))
        self.list_styles = [el for el in list_styles if el is not None]
        self.body_prs = [el for el in (_tx_child(e, "a:bodyPr") for e in (element, layout_ph, master_ph)) if el is not None]
        self.is_title = ph_type in _TITLE_TYPES
        self.major, self.minor = _theme_fonts(master)

    def levels(self, paragraph_el) -> list:
        """The `a:lvlNpPr` of every list style for the paragraph's level, in priority order."""
        p_pr = paragraph_el.find(qn("a:pPr"))
        level = min(max((_int(p_pr.get("lvl")) if p_pr is not None else None) or 0, 0), 8)
        tag = qn(f"a:lvl{level + 1}pPr")
        return [el for el in (styles.find(tag) for styles in self.list_styles) if el is not None]

    def typeface(self, typeface: str | None) -> str:
        if not typeface:
            typeface = "+mj-lt" if self.is_title else "+mn-lt"
        if typeface.startswith("+mj-"):
            return self.major or typeface
        if typeface.startswith("+mn-"):
            return self.minor or typeface
        return typeface

    def body(self) -> _Body:
        insets = {
            name: next((v for b in self.body_prs if (v := _int(b.get(name))) is not None), default)
            for name, default in _DEFAULT_INSETS.items()
        }
        wrap = next((b.get("wrap") for b in self.body_prs if b.get("wrap")), "square")
        fit = next((child for b in self.body_prs for child in b if child.tag in _AUTOFIT), None)
        scale = reduction = None
        if fit is not None and fit.tag == qn("a:normAutofit"):
            scale, reduction = _percent(fit.get("fontScale")), _percent(fit.get("lnSpcReduction"))
        return _Body(
            insets=insets,
            wrap=wrap != "none",
            autofit=_AUTOFIT[fit.tag] if fit is not None else "none",
            font_scale=scale if scale is not None else 1.0,
            line_spacing_reduction=reduction if reduction is not None else 0.0,
        )


def _bool(value: str | None) -> bool:
    return value in ("1", "true")


def _run_properties(run_el):
    """`a:rPr` of a run, field or break; an rPr-like element (such as `a:endParaRPr`) is used as is."""
    if run_el is None:
        return None
    if run_el.tag in _RUN_TAGS:
        return run_el.find(qn("a:rPr"))
    return run_el


def _resolve(run_el, paragraph_el, inheritance: _Inheritance) -> RunStyle:
    p_pr = paragraph_el.find(qn("a:pPr"))
    sources = [_run_properties(run_el), p_pr.find(qn("a:defRPr")) if p_pr is not None else None]
    sources += [level.find(qn("a:defRPr")) for level in inheritance.levels(paragraph_el)]
    sources = [el for el in sources if el is not None]

    def first(name):
        return next((el.get(name) for el in sources if el.get(name) is not None), None)

    latin = (el.find(qn("a:latin")) for el in sources)
    typeface = next((el.get("typeface") for el in latin if el is not None and el.get("typeface")), None)
    size = next((v for el in sources if (v := _int(el.get("sz"))) is not None), None)
    return RunStyle(
        family=inheritance.typeface(typeface),
        size_pt=size / 100 if size is not None else DEFAULT_SIZE_PT,
        bold=_bool(first("b")),
        italic=_bool(first("i")),
    )


def effective_run_style(run_el, paragraph_el, shape, slide) -> RunStyle:
    """Family, size and weight a run is shown with, resolved through PowerPoint's inheritance chain.

    `run_el` is an `a:r`, `a:fld` or `a:br` element (or an `a:endParaRPr`); `paragraph_el` its `a:p`.
    """
    return _resolve(run_el, paragraph_el, _Inheritance(shape, slide))


# --- font lookup -------------------------------------------------------------------------------


@dataclass(frozen=True)
class _Face:
    path: Path
    index: int
    style: str


def _faces(path: Path):
    """(index, family, style) of every face in a font file; collections hold several."""
    index = 0
    while True:
        try:
            family, style = ImageFont.truetype(str(path), index=index).getname()
        except (OSError, ValueError):
            return
        if family:
            yield index, family, style or ""
        if path.suffix.lower() not in _COLLECTION_SUFFIXES:
            return
        index += 1


@lru_cache(maxsize=None)
def _font_index(dirs: tuple[str, ...]) -> tuple[dict[str, list[_Face]], dict[str, list[_Face]]]:
    """Faces by family name and by full name ("family style"), in deterministic scan order.

    Folders that do not exist are skipped.
    """
    families: dict[str, list[_Face]] = {}
    full_names: dict[str, list[_Face]] = {}
    seen: set[Path] = set()
    for directory in dirs:
        for root, subdirs, files in os.walk(Path(directory).expanduser()):
            subdirs.sort()
            for name in sorted(files):
                path = Path(root, name)
                if path.suffix.lower() not in _FONT_SUFFIXES or not path.is_file():
                    continue
                real = path.resolve()
                if real in seen:
                    continue
                seen.add(real)
                for index, family, style in _faces(path):
                    face = _Face(path, index, style)
                    families.setdefault(family.casefold(), []).append(face)
                    if style:
                        full_names.setdefault(f"{family} {style}".casefold(), []).append(face)
    return families, full_names


def _style_rank(style: str, bold: bool, italic: bool) -> int:
    """0 exact style, 1 matching weight/slant with extra words, 2–3 the plain face, 4 anything else."""
    words = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", style).casefold().replace("-", " ").split()
    face_bold = "bold" in words
    face_italic = "italic" in words or "oblique" in words
    canonical = all(word in _RIBBI_WORDS for word in words)
    if (face_bold, face_italic) == (bold, italic):
        return 0 if canonical else 1
    if not (face_bold or face_italic):
        return 2 if canonical else 3
    return 4


def _find_face(family: str, bold: bool, italic: bool) -> _Face | None:
    families, full_names = _font_index(tuple(_FONT_DIRS))
    key = family.casefold()
    faces = families.get(key) or full_names.get(key)
    if not faces:
        return None
    return min(faces, key=lambda face: _style_rank(face.style, bold, italic))


def find_font(family: str, bold: bool = False, italic: bool = False) -> Path | None:
    """The installed font file for `family`, matched by the name inside the file.

    Prefers the face with the requested weight and slant and falls back to the family's plain face;
    None when the family is not installed. The font folders are scanned once per process.
    """
    face = _find_face(family, bold, italic)
    return face.path if face is not None else None


@lru_cache(maxsize=256)
def _load_font(path: Path, index: int, size: float) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(path), size, index=index)


@lru_cache(maxsize=64)
def _fallback_font(size: float) -> ImageFont.FreeTypeFont:
    return ImageFont.load_default(size)


@dataclass(frozen=True)
class _Metrics:
    """Single-line extent of a font at a size, in pt; `missing` names a family measured by the fallback."""

    ascent: float
    descent: float
    missing: str | None = None


def _font_for(style: RunStyle) -> tuple[ImageFont.FreeTypeFont, _Metrics]:
    """The font that measures widths for `style`, with its ascent and descent.

    A family that is not installed, or whose file can no longer be read, is measured with Pillow's
    default font and a single line of 1.2 × size.
    """
    size_pt = max(style.size_pt, 1.0)
    face = _find_face(style.family, style.bold, style.italic)
    if face is not None:
        try:
            font = _load_font(face.path, face.index, size_pt * _SCALE)
        except (OSError, ValueError):
            pass
        else:
            ascent, descent = font.getmetrics()
            return font, _Metrics(ascent / _SCALE, descent / _SCALE)
    metrics = _Metrics(size_pt, (FALLBACK_LINE_HEIGHT - 1) * size_pt, style.family)
    return _fallback_font(size_pt * _SCALE), metrics


# --- measurement -------------------------------------------------------------------------------

# (kind, width in pt, ascent, descent, family missing for a break) with kind "word" | "space" | "break"
_Token = tuple[str, float, float, float, str | None]


def _first_child(sources: list, tag: str):
    for element in sources:
        found = element.find(qn(tag))
        if found is not None:
            return found
    return None


def _spacing(element) -> tuple[str, float] | None:
    """("pct", fraction of a line) or ("pts", points) from an `a:lnSpc`/`a:spcBef`/`a:spcAft`."""
    if element is None:
        return None
    pct, pts = element.find(qn("a:spcPct")), element.find(qn("a:spcPts"))
    value = _percent(pct.get("val")) if pct is not None else None
    if value is not None:
        return "pct", value
    value = _int(pts.get("val")) if pts is not None else None
    if value is not None:
        return "pts", value / 100
    return None


def _spacing_pt(spacing: tuple[str, float] | None, line_pt: float) -> float:
    if spacing is None:
        return 0.0
    kind, value = spacing
    return value * line_pt if kind == "pct" else value


def _tokens(paragraph_el, inheritance: _Inheritance, missing: set[str]) -> list[_Token]:
    """The paragraph as words, spaces and breaks; a word split across runs is one token."""
    tokens: list[_Token] = []
    for child in paragraph_el:
        if child.tag not in _RUN_TAGS:
            continue
        font, metrics = _font_for(_resolve(child, paragraph_el, inheritance))
        if child.tag == qn("a:br"):
            tokens.append(("break", 0.0, metrics.ascent, metrics.descent, metrics.missing))
            continue
        text = child.findtext(qn("a:t")) or ""
        if text and metrics.missing:
            missing.add(metrics.missing)
        for i, part in enumerate(_BREAK.split(text)):
            if i:
                tokens.append(("break", 0.0, metrics.ascent, metrics.descent, None))
            for chunk in _TOKEN.findall(part):
                width = font.getlength(chunk) / _SCALE
                if chunk.isspace():
                    tokens.append(("space", width, metrics.ascent, metrics.descent, None))
                elif tokens and tokens[-1][0] == "word":
                    _, previous, ascent, descent, _ = tokens.pop()
                    tokens.append(("word", previous + width, max(ascent, metrics.ascent), max(descent, metrics.descent), None))
                else:
                    tokens.append(("word", width, metrics.ascent, metrics.descent, None))
    return tokens


def _wrap(tokens: list[_Token], available_pt: float | None, end_metrics, missing: set[str]) -> tuple[list[float], float]:
    """Greedy word wrap; returns every line's single height (largest ascent + largest descent) and
    the widest word, in pt.

    `available_pt` None means no wrapping. A word wider than the line is broken across lines, as
    PowerPoint does. A line holding nothing but a break takes the break's metrics, an empty last
    line those of the paragraph end (`end_metrics()`).
    """
    lines: list[float] = []
    line: tuple[float, float] | None = None  # (ascent, descent) of the current line
    width = widest = 0.0
    has_word = False
    for kind, token_width, ascent, descent, token_missing in tokens:
        if kind == "break":
            if line is None:
                line = (ascent, descent)
                if token_missing:
                    missing.add(token_missing)
            lines.append(sum(line))
            line, width, has_word = None, 0.0, False
            continue
        is_word = kind == "word"
        if is_word:
            widest = max(widest, token_width)
            if available_pt is not None and has_word and width + token_width > available_pt:
                lines.append(sum(line))
                line, width, has_word = None, 0.0, False
        line = (ascent, descent) if line is None else (max(line[0], ascent), max(line[1], descent))
        if is_word and available_pt and not has_word and width + token_width > available_pt:
            extra = math.ceil((width + token_width) / available_pt) - 1
            lines.extend([sum(line)] * extra)
            width -= extra * available_pt
        has_word = has_word or is_word
        width += token_width
    if line is None:
        end = end_metrics()
        if end.missing:
            missing.add(end.missing)
        line = (end.ascent, end.descent)
    lines.append(sum(line))
    return lines, widest


def measure_paragraphs(shape, slide) -> Measurement:
    """Lines per paragraph, text height and widest word of a shape's text, wrapped like PowerPoint.

    Words wrap greedily at the shape width minus the left/right insets and the paragraph's left
    margin. A line is as high as its fonts' ascent + descent (1.2 × size for missing fonts) times
    the line spacing. Spacing between paragraphs counts; space before the first and after the last
    paragraph does not. Fonts that are not installed are measured with Pillow's default font and
    listed in `fonts_missing`.
    """
    tx_body = shape._element.find(qn("p:txBody"))
    if tx_body is None:
        return Measurement([], 0.0, 0.0, set())
    inheritance = _Inheritance(shape, slide)
    body = inheritance.body()
    text_width = absolute_box(shape).w * EMU_PER_CM - body.insets["lIns"] - body.insets["rIns"]
    missing: set[str] = set()
    lines_per_paragraph: list[int] = []
    height_pt = widest_pt = 0.0
    paragraphs = tx_body.findall(qn("a:p"))
    for number, paragraph in enumerate(paragraphs):
        p_pr = paragraph.find(qn("a:pPr"))
        sources = ([p_pr] if p_pr is not None else []) + inheritance.levels(paragraph)
        margin = next((v for el in sources if (v := _int(el.get("marL"))) is not None), 0)
        available = max(text_width - margin, 0) / EMU_PER_PT if body.wrap else None
        end = paragraph.find(qn("a:endParaRPr"))
        lines, widest = _wrap(
            _tokens(paragraph, inheritance, missing),
            available,
            lambda: _font_for(_resolve(end, paragraph, inheritance))[1],
            missing,
        )
        kind, value = _spacing(_first_child(sources, "a:lnSpc")) or ("pct", 1.0)
        height_pt += sum(line * value for line in lines) if kind == "pct" else value * len(lines)
        if number > 0:
            height_pt += _spacing_pt(_spacing(_first_child(sources, "a:spcBef")), lines[0])
        if number < len(paragraphs) - 1:
            height_pt += _spacing_pt(_spacing(_first_child(sources, "a:spcAft")), lines[-1])
        lines_per_paragraph.append(len(lines))
        widest_pt = max(widest_pt, widest)
    to_cm = EMU_PER_PT / EMU_PER_CM
    height_cm = (body.insets["tIns"] + body.insets["bIns"]) / EMU_PER_CM + height_pt * to_cm
    return Measurement(
        lines_per_paragraph,
        round(height_cm, 2),
        round(widest_pt * to_cm, 2),
        missing,
        autofit=body.autofit,
        font_scale=body.font_scale,
        line_spacing_reduction=body.line_spacing_reduction,
    )
