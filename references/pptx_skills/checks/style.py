"""Style rules: theme colours and fonts instead of hard-coded ones, alt text for pictures."""

from pptx.oxml.ns import qn

from ..fonts import _theme_fonts, find_font
from . import rule

_OWNERS = tuple(qn(t) for t in ("p:sp", "p:pic", "p:cxnSp", "p:grpSp", "p:graphicFrame"))


def _owner_id(element) -> int | None:
    """Id of the shape that contains `element`."""
    for ancestor in element.iterancestors():
        if ancestor.tag in _OWNERS:
            found = ancestor.xpath("./*[1]/p:cNvPr")
            return int(found[0].get("id")) if found else None
    return None


def _by_id(slide) -> dict:
    return {item["id"]: item for item, _ in slide.shapes()}


def _tree(slide):
    return slide.pptx._element.find(qn("p:cSld")).find(qn("p:spTree"))


@rule("rgb-color", "style", "warning")
def rgb_color(ctx, slide):
    shapes = _by_id(slide)
    counts: dict[int, list[str]] = {}
    for el in _tree(slide).iter(qn("a:srgbClr")):
        if any(a.tag == qn("a:blip") for a in el.iterancestors()):
            continue
        owner = _owner_id(el)
        if owner in shapes:
            counts.setdefault(owner, []).append(el.get("val", ""))
    return [
        ctx.finding(slide, shapes[sid], f"the shape uses fixed RGB colours ({', '.join(sorted(set(vals)))})", "use the theme colour that matches (accent1…6, tx1, bg1 …) so the deck follows the master")
        for sid, vals in counts.items()
    ]


@rule("foreign-font", "style", "warning")
def foreign_font(ctx, slide):
    theme = {f for f in _theme_fonts(ctx.prs.slide_master) if f}
    shapes = _by_id(slide)
    found: dict[int, set[str]] = {}
    for el in _tree(slide).iter(qn("a:latin")):
        face = el.get("typeface")
        if not face or face.startswith("+") or face in theme:
            continue
        owner = _owner_id(el)
        if owner in shapes:
            found.setdefault(owner, set()).add(face)
    return [
        ctx.finding(slide, shapes[sid], f"the text uses a font outside the theme ({', '.join(sorted(faces))})", f"remove the font override so the text uses the theme fonts ({' / '.join(sorted(theme)) or 'theme fonts'})")
        for sid, faces in found.items()
    ]


def missing_theme_fonts(prs) -> list[str]:
    """The theme fonts of the deck that are not installed on this machine."""
    names = [f for f in dict.fromkeys(_theme_fonts(prs.slide_master)) if f and not f.startswith("+")]
    return [f for f in names if find_font(f) is None]


@rule("missing-font", "style", "warning", scope="deck")
def missing_font(ctx, _):
    missing = missing_theme_fonts(ctx.prs)
    if not missing:
        return []
    return [ctx.finding(None, None, f"the theme font {', '.join(missing)} is not installed, so measured text sizes are approximations", "install the font (the layout checks and renders then match PowerPoint) or accept the approximate results")]


@rule("alt-text", "style", "info")
def alt_text(ctx, slide):
    return [
        ctx.finding(slide, item, "the picture has no alt text", "describe the picture in its alt text (or mark it decorative in PowerPoint)")
        for item, _ in slide.shapes()
        if item["kind"] == "picture" and not item["alt"].strip()
    ]
