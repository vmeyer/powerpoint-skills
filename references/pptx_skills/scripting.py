"""`pptx exec`: run an agent-written python-pptx script against a working copy, with the safety of apply.

The script runs in this process with the globals of `_globals`. The tool saves; the script never does. Nothing
is written unless the script finishes, the slides are normalised (ids, sections, part names, slide numbers) and
the structure rules pass; then the deck is saved atomically, with snapshot and the script in the change log.
"""

import copy
import hashlib
import traceback
from pathlib import Path

from lxml import etree
from pptx.dml.color import ColorFormat
from pptx.dml.fill import FillFormat
from pptx.dml.line import LineFormat
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

from . import cli, slides
from .changeset import Context, finish, open_working_copy, result_lines, slug_label
from .checks import structure_errors
from .colors import apply_color, parse_color
from .errors import PptxSkillsError

_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def theme_color(target, color: str) -> None:
    """Color a fill (solid), a line, a font or any color format (e.g. a gradient stop's `color`) with a theme
    color such as ``accent1+40``, a ``#RRGGBB`` or (fills and lines) ``none``."""
    ref = parse_color(str(color))
    if isinstance(target, FillFormat):
        if ref is None:
            target.background()
            return
        target.solid()
        target = target.fore_color
    elif isinstance(target, LineFormat):
        if ref is None:
            target.fill.background()
            return
        target = target.color
    elif not isinstance(target, ColorFormat) and isinstance(getattr(target, "color", None), ColorFormat):
        target = target.color  # a font
    if not isinstance(target, ColorFormat):
        raise PptxSkillsError(f"theme_color takes a fill, line, font or color format, not {type(target).__name__}")
    if ref is None:
        raise PptxSkillsError('"none" only applies to a fill or a line')
    apply_color(target, ref)


def plain(shape):
    """Drop the theme style python-pptx gives a new auto shape (accent fill, outline, effects, light text), so it
    starts as an added shape of a change set: no fill, no line, the default text color. Returns the shape."""
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)
    return shape


def _globals(ctx: Context, script: Path) -> dict:
    prs = ctx.prs

    def slide_by_id(slide_id: int):
        return slides.resolve_slide(prs, f"id:{slide_id}")

    def layout(name: str):
        return slides.layout_named(prs, name)

    def add_slide(layout_name: str, after=None):
        """A slide with the layout's footer, date and number, a new id, in the right section; at the end or after
        `after` (a slide, a 1-based number or "id:<n>")."""
        anchor = after if after is None or hasattr(after, "slide_id") else slides.resolve_slide(prs, after)
        slide = slides.add_slide(prs, layout_name, anchor, ctx.max_slide_id, ctx.messages)
        ctx.max_slide_id = max(ctx.max_slide_id, slide.slide_id)
        return slide

    return {
        "__name__": "__pptx_exec__",
        "__file__": str(script),
        "prs": prs,
        "deck_dir": ctx.deck_dir,
        "brief": ctx.brief,
        "cm": Cm,
        "pt": Pt,
        "theme_color": theme_color,
        "plain": plain,
        "slide_by_id": slide_by_id,
        "layout": layout,
        "add_slide": add_slide,
    }


def _fingerprint(slide) -> str:
    """A hash of the slide's XML and its charts; slide-number fields count without their cached number."""
    element = copy.deepcopy(slide._element)
    for field in element.iter(qn("a:fld")):
        text = field.find(qn("a:t"))
        if field.get("type") == "slidenum" and text is not None:
            text.text = ""
    digest = hashlib.sha1(etree.tostring(element))
    for rel in slide.part.rels.values():
        if rel.reltype == RT.CHART and not rel.is_external:
            digest.update(rel.target_part.blob)
    return digest.hexdigest()


def _fingerprints(prs) -> dict[int, str]:
    return {s.slide_id: _fingerprint(s) for s in prs.slides}


def _in_a_section(prs, slide_id: int) -> bool:
    return any(ref.get("id") == str(slide_id) for ref in prs.part._element.iter(f"{{{_P14}}}sldId"))


def _normalise(ctx: Context, before: dict, floor: int) -> None:
    """What a script that used python-pptx directly may have left out: new slide ids above every id the deck folder
    used (R26), a section for each new slide (R12), unique slide part names and current slide numbers (R19).

    `before` maps the slide ids before the script to their slide parts; `floor` is the highest id used before."""
    prs = ctx.prs
    id_list = list(prs.slides._sldIdLst)
    has_sections = next(prs.part._element.iter(f"{{{_P14}}}sectionLst"), None) is not None
    for index, sld_id in enumerate(id_list):
        if before.get(sld_id.id) is prs.part.related_part(sld_id.rId):
            continue  # a slide that was there before
        if sld_id.id <= floor:
            note = slides.allocate_slide_id(prs, sld_id, ctx.max_slide_id)
            if note:
                ctx.messages.append(note)
        ctx.max_slide_id = max(ctx.max_slide_id, sld_id.id)
        if has_sections and not _in_a_section(prs, sld_id.id):
            slides.place_in_section(prs, sld_id.id, id_list[index - 1].id if index else None)
    prs.part.rename_slide_parts([sld_id.rId for sld_id in prs.slides._sldIdLst])
    slides.refresh_slide_numbers(prs)


def _failure(script: Path, exc: BaseException) -> PptxSkillsError:
    if isinstance(exc, SyntaxError):
        return PptxSkillsError(f"{script.name} line {exc.lineno}: SyntaxError: {exc.msg}; nothing was changed")
    lines = [f.lineno for f in traceback.extract_tb(exc.__traceback__) if f.filename == str(script)]
    where = f" line {lines[-1]}" if lines else ""
    what = str(exc) if isinstance(exc, PptxSkillsError) else f"{type(exc).__name__}: {exc}"
    return PptxSkillsError(f"{script.name}{where}: {what}; nothing was changed")


def run_script(deck_path: Path, script: Path, outside_deck: bool = False) -> dict:
    """Run the python-pptx `script` against the working copy `deck_path` (any .pptx with `outside_deck`) and save
    the result like apply: all or nothing, with snapshot and the script logged as `changes/NNN-<name>.py`.

    Touched are the slides whose XML changed and the new ones. Returns the result of `finish`.
    """
    script = Path(script)
    try:
        source = script.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as exc:
        raise PptxSkillsError(f"cannot read the script {script}: {exc}") from exc
    try:
        code = compile(source, str(script), "exec")
    except SyntaxError as exc:
        raise _failure(script, exc) from exc
    ctx = open_working_copy(deck_path, outside_deck, command="exec")
    prs = ctx.prs
    before, floor = _fingerprints(prs), ctx.max_slide_id
    parts = {s.slide_id: s.part for s in prs.slides}

    def refuse_save(*_args, **_kwargs):
        raise PptxSkillsError("the script must not save: pptx exec saves the deck when the script has finished")

    prs.save = refuse_save  # shadows the method for the script's run only
    try:
        exec(code, _globals(ctx, script))  # noqa: S102 - running the agent's script is this command's job
    except KeyboardInterrupt:
        raise
    except BaseException as exc:  # SystemExit too: nothing is saved unless the script runs to its end
        raise _failure(script, exc) from exc
    finally:
        del prs.save
    _normalise(ctx, parts, floor)
    broken = structure_errors(prs)
    if broken:
        raise PptxSkillsError(f"{script.name} left the deck broken, so nothing was changed: " + "; ".join(broken))
    after = _fingerprints(prs)
    ctx.touched = {slide_id for slide_id, digest in after.items() if before.get(slide_id) != digest}
    return finish(ctx, source, slug_label(script.stem, "script"))


def _handle(args) -> int:
    script = Path(args.script)
    result = run_script(Path(args.deck), script, args.outside_deck)
    head = f"Ran {script.name}" + (f" as change {result['seq']:03d}" if result["seq"] else "")
    cli.emit(args, result, "\n".join([head, *result_lines(result)]))
    return cli.OK


def register(subparsers) -> None:
    parser = subparsers.add_parser(
        "exec", help="Run a python-pptx script against a working copy: snapshot, change log, all or nothing."
    )
    parser.add_argument("deck", help="working copy (.pptx)")
    parser.add_argument("script", help="python-pptx script (.py); globals: prs, deck_dir, brief, cm, pt, theme_color, "
                        "plain, slide_by_id, layout, add_slide")
    parser.add_argument(
        "--outside-deck", action="store_true",
        help="also change a .pptx that is not a deck folder's working copy, in place and without snapshot or change log",
    )
    parser.add_argument("--json", action="store_true", help="machine-readable output")
    parser.set_defaults(handler=_handle)
