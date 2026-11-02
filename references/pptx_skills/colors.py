"""Theme color references such as ``accent1+60``: parse, format, apply and read."""

import re
from dataclasses import dataclass

from pptx.dml.color import ColorFormat, RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.ns import qn

from pptx_skills.errors import PptxSkillsError

ROLE_TO_MSO: dict[str, MSO_THEME_COLOR] = {
    "dk1": MSO_THEME_COLOR.DARK_1,
    "lt1": MSO_THEME_COLOR.LIGHT_1,
    "dk2": MSO_THEME_COLOR.DARK_2,
    "lt2": MSO_THEME_COLOR.LIGHT_2,
    "tx1": MSO_THEME_COLOR.TEXT_1,
    "tx2": MSO_THEME_COLOR.TEXT_2,
    "bg1": MSO_THEME_COLOR.BACKGROUND_1,
    "bg2": MSO_THEME_COLOR.BACKGROUND_2,
    **{f"accent{i}": getattr(MSO_THEME_COLOR, f"ACCENT_{i}") for i in range(1, 7)},
    "hlink": MSO_THEME_COLOR.HYPERLINK,
    "folHlink": MSO_THEME_COLOR.FOLLOWED_HYPERLINK,
}

_ROLE_RE = "|".join(sorted(ROLE_TO_MSO, key=len, reverse=True))
_THEME_PATTERN = re.compile(rf"^({_ROLE_RE})(?:([+-])(\d{{1,2}}))?$")
_RGB_PATTERN = re.compile(r"^#([0-9A-Fa-f]{6})$")


@dataclass(frozen=True)
class ColorRef:
    role: str | None
    rgb: str | None
    brightness: int = 0


def parse_color(text: str) -> ColorRef | None:
    """Parse ``accent1``, ``accent1+60``, ``dk2-25``, ``#RRGGBB`` or ``none``."""
    text = text.strip()
    if text == "none":
        return None
    if text.startswith("#"):
        match = _RGB_PATTERN.match(text)
        if not match:
            raise PptxSkillsError(f"unknown color {text!r}; use #RRGGBB or a theme role such as accent1 or tx1-25")
        return ColorRef(None, match.group(1).upper(), 0)
    match = _THEME_PATTERN.match(text)
    if not match or (match.group(3) is not None and not 1 <= int(match.group(3)) <= 99):
        raise PptxSkillsError(f"unknown color {text!r}")
    brightness = 0
    if match.group(2):
        brightness = int(match.group(3)) * (1 if match.group(2) == "+" else -1)
    return ColorRef(match.group(1), None, brightness)


def format_color(ref: ColorRef) -> str:
    if ref.rgb is not None:
        return f"#{ref.rgb}"
    return f"{ref.role}{ref.brightness:+d}" if ref.brightness else str(ref.role)


def apply_color(color_format: ColorFormat, ref: ColorRef) -> None:
    """Set the color exactly as `ref` says: tints, shades and saturation the color carried before (e.g. from
    python-pptx's default gradient) are dropped; only its transparency (`a:alpha`) stays."""
    color = getattr(color_format._color, "_xClr", None)
    alpha = []
    if color is not None:
        alpha = [child for child in color if child.tag == qn("a:alpha")]
        for child in list(color):
            if child.tag != qn("a:alpha"):
                color.remove(child)
    if ref.rgb is not None:
        color_format.rgb = RGBColor.from_string(ref.rgb)
    else:
        color_format.theme_color = ROLE_TO_MSO[ref.role]
        color_format.brightness = ref.brightness / 100
    new = color_format._color._xClr
    if new is not color:  # python-pptx replaced the element (theme color <-> RGB): carry the transparency over
        for child in reversed(alpha):
            new.insert(0, child)


_SYSTEM_COLORS = {"windowText": "#000000", "window": "#FFFFFF"}  # a:sysClr without a cached lastClr


def _val(el, name: str) -> int | None:
    child = el.find(qn(f"a:{name}"))
    try:
        return int(child.get("val")) if child is not None else None
    except (TypeError, ValueError):  # no val, or not a whole number
        return None


def read_color(fill_el) -> str | None:
    """Read the color held by a fill-like element (``a:solidFill``, ``a:fgClr``, ...); None when there is none.

    A brightness that rounds to 0 % is left out (``accent1``, never ``accent1+0``).
    """
    scheme = fill_el.find(qn("a:schemeClr"))
    if scheme is not None:
        role = scheme.get("val")
        if not role:
            return None
        lum_off, lum_mod = _val(scheme, "lumOff"), _val(scheme, "lumMod")
        shift = 0
        if lum_off is not None:
            shift = round(lum_off / 1000)
        elif lum_mod is not None and lum_mod < 100000:
            shift = -round(100 - lum_mod / 1000)
        return f"{role}{shift:+d}" if shift else role
    srgb = fill_el.find(qn("a:srgbClr"))
    if srgb is not None:
        return f"#{srgb.get('val').upper()}" if srgb.get("val") else None
    sys_clr = fill_el.find(qn("a:sysClr"))
    if sys_clr is not None:
        last = sys_clr.get("lastClr")
        return f"#{last.upper()}" if last else _SYSTEM_COLORS.get(sys_clr.get("val"))
    return None
