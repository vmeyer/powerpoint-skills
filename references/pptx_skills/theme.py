"""`pptx theme`: show or set the theme colors and fonts of a working copy (snapshot and change log as for apply)."""

import re
from pathlib import Path

from pptx.oxml.ns import qn

from . import cli
from .brief import _theme
from .changeset import finish, open_working_copy
from .deckdir import load_deck, refresh_brief
from .errors import PptxSkillsError
from .fonts import find_font
from .neutral import THEME_ROLES, set_theme
from .opc import open_presentation

_HEX = re.compile(r"^#[0-9A-Fa-f]{6}$")
# the color map's names for text and background; PowerPoint's default map is tx1=dk1, bg1=lt1, tx2=dk2, bg2=lt2
_MAPPED = {"tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2"}


def _deck_file(path: Path) -> Path:
    """A deck folder's working copy, or the file itself."""
    path = Path(path)
    return path / load_deck(path)["working_copy"] if path.is_dir() else path


def _theme_role(prs, role: str) -> str:
    """The theme color behind `role`: tx1, bg1, tx2 and bg2 through the master's color map."""
    if role in THEME_ROLES:
        return role
    if role in _MAPPED:
        color_map = prs.slide_master._element.find(qn("p:clrMap"))
        mapped = color_map.get(role) if color_map is not None else None
        return mapped if mapped in THEME_ROLES else _MAPPED[role]
    raise PptxSkillsError(f"unknown theme color {role!r}; roles: {', '.join(THEME_ROLES)}, or tx1, bg1, tx2, bg2")


def parse_options(colors: list[str], fonts: list[str]) -> tuple[dict, dict]:
    """`role=#RRGGBB` and `major=<family>` / `minor=<family>` options as two dicts; roles are checked later."""
    parsed_colors, parsed_fonts = {}, {}
    for item in colors:
        role, sep, value = item.partition("=")
        if not sep or not role.strip():
            raise PptxSkillsError(f"--color expects role=#RRGGBB, e.g. accent1=#0B5CAD; got {item!r}")
        if not _HEX.match(value.strip()):
            raise PptxSkillsError(f"--color {role.strip()} needs a color as #RRGGBB, e.g. #0B5CAD; got {value!r}")
        parsed_colors[role.strip()] = value.strip().upper()
    for item in fonts:
        key, sep, family = item.partition("=")
        if not sep or key.strip() not in ("major", "minor") or not family.strip():
            raise PptxSkillsError(f"--font takes major or minor, as major=<family> (headings) or minor=<family> (body); got {item!r}")
        parsed_fonts[key.strip()] = family.strip()
    return parsed_colors, parsed_fonts


def apply_theme(deck_path: Path, colors: dict, fonts: dict) -> dict:
    """Set theme colors ({role: "#RRGGBB"}, tx1/bg1/tx2/bg2 allowed) and fonts ({"major"|"minor": family}) in a
    working copy, all or nothing, and refresh its style brief. Returns the result of `finish` plus the theme."""
    bad = [value for value in colors.values() if not _HEX.match(value)]
    if bad:
        raise PptxSkillsError(f"theme colors are #RRGGBB values; got {bad[0]!r}")
    ctx = open_working_copy(_deck_file(deck_path), command="theme")
    resolved = {_theme_role(ctx.prs, role): value.upper() for role, value in colors.items()}
    set_theme(ctx.prs, resolved, fonts)
    result = finish(ctx, {"label": "theme", "theme": {"colors": resolved, "fonts": fonts}}, "theme")
    if ctx.deck_dir is not None:
        refresh_brief(ctx.deck_dir, load_deck(ctx.deck_dir))
    result["messages"].append(
        f"every slide now looks different and earlier renders are stale; re-render with `pptx render {ctx.deck_path} --sheet`"
    )
    for family in dict.fromkeys(fonts.values()):
        if find_font(family) is None:
            result["messages"].append(
                f"the font {family} is not installed on this machine: renders and text measurement use a fallback, "
                "and whoever opens the deck needs the font too"
            )
    return {**result, **_theme(ctx.prs)}


def _theme_lines(theme: dict) -> list[str]:
    fonts = theme["fonts"]
    return [
        "colors: " + ", ".join(f"{role} {value}" for role, value in theme["colors"].items()),
        f"fonts: {fonts['major']} (headings), {fonts['minor']} (body)",
    ]


def _handle(args) -> int:
    colors, fonts = parse_options(args.color or [], args.font or [])
    if not colors and not fonts:
        path = _deck_file(Path(args.deck))
        theme = _theme(open_presentation(path))
        cli.emit(args, theme, "\n".join([f"Theme of {path}:", *_theme_lines(theme)]))
        return cli.OK
    result = apply_theme(Path(args.deck), colors, fonts)
    head = "Theme set" + (f" as change {result['seq']:03d}" if result["seq"] else "")
    cli.emit(args, result, "\n".join([head, *_theme_lines(result), *(f"note: {m}" for m in result["messages"])]))
    return cli.OK


def register(subparsers) -> None:
    parser = subparsers.add_parser(
        "theme", help="Show or set the theme colors and fonts of a working copy (snapshot and change log as for apply)."
    )
    parser.add_argument("deck", help="deck folder or its working copy (.pptx)")
    parser.add_argument("--color", action="append", metavar="ROLE=#RRGGBB",
                        help="a theme color: dk1 lt1 dk2 lt2 accent1..accent6 hlink folHlink, or tx1 bg1 tx2 bg2; repeatable")
    parser.add_argument("--font", action="append", metavar="major|minor=FAMILY",
                        help="the heading (major) or body (minor) font; repeatable")
    parser.add_argument("--json", action="store_true", help="machine-readable output")
    parser.set_defaults(handler=_handle)
