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


def parse_color(text: str, allow_rgb: bool = False) -> ColorRef | None:
    """Parse ``accent1``, ``accent1+60``, ``dk2-25``, ``#RRGGBB`` (if allowed) or ``none``."""
    text = text.strip()
    if text == "none":
        return None
    if text.startswith("#"):
        match = _RGB_PATTERN.match(text)
        if not match:
            raise PptxSkillsError(f"unknown color {text!r}")
        if not allow_rgb:
            raise PptxSkillsError(
                f"raw color {text!r} is not allowed; use a theme role or set allow_rgb"
            )
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
    if ref.rgb is not None:
        color_format.rgb = RGBColor.from_string(ref.rgb)
        return
    color_format.theme_color = ROLE_TO_MSO[ref.role]
    color_format.brightness = ref.brightness / 100


def _val(el, name: str) -> int | None:
    child = el.find(qn(f"a:{name}"))
    return int(child.get("val")) if child is not None else None


def read_color(fill_el) -> str | None:
    """Read the color held by a fill-like element (``a:solidFill``, ``a:fgClr``, ...)."""
    scheme = fill_el.find(qn("a:schemeClr"))
    if scheme is not None:
        text = scheme.get("val")
        lum_off, lum_mod = _val(scheme, "lumOff"), _val(scheme, "lumMod")
        if lum_off is not None:
            return f"{text}+{round(lum_off / 1000)}"
        if lum_mod is not None and lum_mod < 100000:
            return f"{text}-{round(100 - lum_mod / 1000)}"
        return text
    srgb = fill_el.find(qn("a:srgbClr"))
    if srgb is not None:
        return f"#{srgb.get('val').upper()}"
    sys_clr = fill_el.find(qn("a:sysClr"))
    if sys_clr is not None:
        return f"#{sys_clr.get('lastClr').upper()}"
    return None
