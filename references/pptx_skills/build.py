"""`pptx build`: one step of work in one command: apply a change set or run a script, check and render what changed."""

import sys
from pathlib import Path

from . import cli
from .changeset import apply_changeset, read_changeset, result_lines
from .checks import format_report, run_checks
from .errors import PptxSkillsError
from .opc import open_presentation
from .render import add_renderer_option, default_out_dir, render, render_slide_sheet
from .scripting import run_script


def build(deck_path: Path, change: Path, strict: bool = False, renderer: str | None = None, dpi: int = 120,
          outside_deck: bool = False) -> dict:
    """Apply `change` (a change set, or a python-pptx script ending in .py) to the working copy, check the touched
    slides (every rule with `strict`) and render them, plus a contact sheet when there are several.

    A failing change raises and leaves everything as it was. A failing render does not undo the change: the
    result then holds `render_error`.
    """
    deck_path, change = Path(deck_path), Path(change)
    if change.suffix.lower() == ".py":
        result = run_script(deck_path, change, outside_deck)
        headline = f"Ran {change.name}"
    else:
        changeset = read_changeset(change)
        result = apply_changeset(deck_path, changeset, outside_deck)
        headline = f"Applied {len(changeset['ops'])} op(s)"
    numbers = {s.slide_id: n for n, s in enumerate(open_presentation(deck_path).slides, 1)}
    touched = sorted(numbers[i] for i in result["touched"] if i in numbers)
    report = run_checks(deck_path, slides=touched, strict=strict)
    notes = [*result["messages"], *report.get("notes", [])]
    data = {
        **result,
        "headline": headline + (f" as change {result['seq']:03d}" if result["seq"] else ""),
        "slides": touched,
        "findings": report["findings"],
        "summary": report["summary"],
        "notes": notes,
        "pngs": [],
        "sheet": None,
    }
    if not touched:
        notes.append("nothing to render: no slide was added or changed")
        return data
    try:
        pngs = render(deck_path, None, touched, dpi, renderer, notes)
        data["pngs"] = [str(p) for p in pngs]
        if len(pngs) > 1:
            data["sheet"] = str(render_slide_sheet(pngs, touched, default_out_dir(deck_path) / "sheet.png"))
    except PptxSkillsError as exc:
        data["render_error"] = str(exc)
    return data


def _text(data: dict) -> str:
    report = {"findings": data["findings"], "summary": data["summary"], "slides": data["slides"]}
    lines = [data["headline"], *result_lines({**data, "messages": data["notes"]}), format_report(report)]
    return "\n".join([*lines, *data["pngs"], *([data["sheet"]] if data["sheet"] else [])])


def _handle(args) -> int:
    data = build(Path(args.deck), Path(args.change), args.strict, args.renderer, args.dpi, args.outside_deck)
    cli.emit(args, data, _text(data))
    if "render_error" in data:
        print(f"error: render failed, the change is saved: {data['render_error']}", file=sys.stderr)
        return cli.ERROR
    return cli.FINDINGS if data["summary"]["error"] else cli.OK


def register(subparsers) -> None:
    parser = subparsers.add_parser(
        "build",
        help="Apply a change set (.json) or run a python-pptx script (.py), then check and render the touched slides.",
    )
    parser.add_argument("deck", help="working copy (.pptx)")
    parser.add_argument("change", help="change set (.json) or python-pptx script (.py)")
    parser.add_argument("--strict", action="store_true", help="check with every rule, not only errors and the writing rules")
    parser.add_argument("--dpi", type=int, default=120, help="render resolution, default 120")
    add_renderer_option(parser)
    parser.add_argument(
        "--outside-deck", action="store_true",
        help="also change a .pptx that is not a deck folder's working copy, in place and without snapshot or change log",
    )
    parser.add_argument("--json", action="store_true", help="machine-readable output")
    parser.set_defaults(handler=_handle)
