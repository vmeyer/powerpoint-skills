"""Style files and the neutral 16:9 template for decks without a PowerPoint master."""

import json
from importlib import resources
from pathlib import Path

from jsonschema import Draft202012Validator
from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

from .errors import PptxSkillsError

NEUTRAL_COLORS = {
    "dk1": "#000000",
    "lt1": "#FFFFFF",
    "dk2": "#333333",
    "lt2": "#EDEDED",
    "accent1": "#1F4E79",
    "accent2": "#2E75B6",
    "accent3": "#7F7F7F",
    "accent4": "#BFBFBF",
    "accent5": "#D9D9D9",
    "accent6": "#C55A11",
    "hlink": "#1F4E79",
    "folHlink": "#7F7F7F",
}
NEUTRAL_FONTS = {"major": "Arial", "minor": "Arial"}
_WIDTH, _HEIGHT = 12192000, 6858000


def validate_style(data, where: str = "style file") -> dict:
    """Check a parsed style file against the schema; errors name the offending key."""
    schema = json.loads(resources.files(__package__).joinpath("schemas/style-file.schema.json").read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise PptxSkillsError(f"{where} must be a JSON object with the keys colors, fonts and rules")
    errors = sorted(Draft202012Validator(schema).iter_errors(data), key=lambda e: list(e.absolute_path))
    if errors:
        error = errors[0]
        path = "/".join(str(p) for p in error.absolute_path)
        if error.validator == "additionalProperties":
            unknown = sorted(k for k in error.instance if k not in error.schema["properties"])
            path = "/".join([*(str(p) for p in error.absolute_path), unknown[0]])
            raise PptxSkillsError(f"{where}: unknown key {path!r}; allowed: {', '.join(error.schema['properties'])}")
        raise PptxSkillsError(f"{where}: {path or '(root)'}: {error.message}")
    return data


def load_style_file(path: Path) -> dict:
    path = Path(path)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise PptxSkillsError(f"cannot read style file {path}: {exc}") from exc
    except ValueError as exc:
        raise PptxSkillsError(f"style file {path} is not valid JSON: {exc}") from exc
    return validate_style(data, f"style file {path}")


def _scale_x(prs) -> None:
    parts = [prs.slide_master, *prs.slide_layouts]
    for part in parts:
        for xfrm in part._element.iter(qn("a:xfrm")):
            for tag, attr in (("a:off", "x"), ("a:ext", "cx"), ("a:chOff", "x"), ("a:chExt", "cx")):
                el = xfrm.find(qn(tag))
                if el is not None and el.get(attr) is not None:
                    el.set(attr, str(round(int(el.get(attr)) * 4 / 3)))


def _style_theme(prs, colors: dict, fonts: dict) -> None:
    part = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(part.blob)
    root.set("name", "Neutral")
    scheme = root.find(".//" + qn("a:clrScheme"))
    scheme.set("name", "Neutral")
    for role, value in colors.items():
        el = scheme.find(qn(f"a:{role}"))
        for child in list(el):
            el.remove(child)
        etree.SubElement(el, qn("a:srgbClr")).set("val", value.lstrip("#").upper())
    font_scheme = root.find(".//" + qn("a:fontScheme"))
    font_scheme.set("name", "Neutral")
    for key, tag in (("major", "a:majorFont"), ("minor", "a:minorFont")):
        font_scheme.find(f"{qn(tag)}/{qn('a:latin')}").set("typeface", fonts[key])
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def neutral_presentation(style: dict | None = None):
    """python-pptx's default template as 16:9 with the neutral (or the style file's) colors and fonts."""
    style = validate_style(style or {})
    prs = Presentation()
    prs.slide_width, prs.slide_height = _WIDTH, _HEIGHT
    _scale_x(prs)
    _style_theme(prs, {**NEUTRAL_COLORS, **style.get("colors", {})}, {**NEUTRAL_FONTS, **style.get("fonts", {})})
    return prs
