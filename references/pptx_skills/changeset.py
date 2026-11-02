"""Apply a validated JSON change set to a working copy: all ops succeed or nothing is written."""

import json
import os
import re
from dataclasses import dataclass, field
from importlib import resources
from pathlib import Path

from jsonschema import Draft202012Validator

from . import cli, slides
from .brief import load_brief
from .deckdir import changed_externally, find_deck_dir, load_deck, record_apply
from .errors import PptxSkillsError
from .geometry import absolute_box
from .opc import lock_file, open_presentation, save_presentation
from .roles import placeholder_role

_TITLE_TYPES = ("TITLE", "CENTER_TITLE", "VERTICAL_TITLE")


@dataclass
class Context:
    prs: object
    deck_path: Path
    deck_dir: Path | None
    brief: dict | None
    examples_path: str | None
    last_slide: object = None
    touched: set[int] = field(default_factory=set)
    messages: list[str] = field(default_factory=list)


OPS: dict = {}


def op(name: str):
    """Register `fn(ctx, spec)` as the handler of the change-set op `name`."""

    def register(fn):
        OPS[name] = fn
        return fn

    return register


def _schema() -> dict:
    text = resources.files(__package__).joinpath("schemas/changeset.schema.json").read_text(encoding="utf-8")
    return json.loads(text)


def _where(error) -> str:
    return "/".join(str(p) for p in error.absolute_path) or "(op)"


def validate_changeset(changeset) -> None:
    """Check the envelope and every op against the schema; errors name the 1-based op number."""
    schema = _schema()
    branches = {b["properties"]["op"]["const"]: b for b in schema["properties"]["ops"]["items"]["oneOf"]}
    if not isinstance(changeset, dict) or not isinstance(changeset.get("ops"), list) or not changeset["ops"]:
        raise PptxSkillsError('a change set is an object {"label": "...", "ops": [ ... ]} with at least one op')
    for number, spec in enumerate(changeset["ops"], 1):
        name = spec.get("op") if isinstance(spec, dict) else None
        if not isinstance(name, str):
            raise PptxSkillsError(f'op {number} (?): every op is an object with an "op" name')
        if name not in branches:
            raise PptxSkillsError(f"op {number} ({name}): unknown op '{name}'; known ops: {', '.join(sorted(branches))}")
        branch = dict(branches[name], **{"$defs": schema["$defs"]})
        errors = sorted(Draft202012Validator(branch).iter_errors(spec), key=lambda e: list(e.absolute_path))
        if errors:
            detail = "; ".join(f"{_where(e)}: {e.message}" for e in errors)
            raise PptxSkillsError(f"op {number} ({name}): {detail}")
    envelope = dict(schema, properties=dict(schema["properties"], ops={"type": "array"}))
    errors = list(Draft202012Validator(envelope).iter_errors(changeset))
    if errors:
        raise PptxSkillsError("; ".join(f"{_where(e)}: {e.message}" for e in errors))


def _role_of(shape, title_box) -> str:
    if shape.is_placeholder:
        return placeholder_role(shape.placeholder_format.type, absolute_box(shape), title_box)
    return {"PICTURE": "picture", "TABLE": "table", "CHART": "chart"}.get(getattr(shape.shape_type, "name", ""), "other")


def _walk(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type is not None and shape.shape_type.name == "GROUP":
            yield from _walk(shape.shapes)


def resolve_shape(slide, target: dict):
    """Find one shape by `{"role"}`, `{"name"}` or `{"id"}`, searching group children too."""
    if not isinstance(target, dict) or len(target) != 1 or next(iter(target)) not in ("role", "name", "id"):
        raise PptxSkillsError('a target is exactly one of {"role": ...}, {"name": ...}, {"id": ...}')
    key, value = next(iter(target.items()))
    shapes = list(_walk(slide.shapes))
    if key == "role":
        title_box = next(
            (absolute_box(s) for s in slide.placeholders if s.placeholder_format.type.name in _TITLE_TYPES), None
        )
        found = [s for s in shapes if _role_of(s, title_box) == value]
    elif key == "name":
        found = [s for s in shapes if s.name == value]
    else:
        found = [s for s in shapes if s.shape_id == value]
    if len(found) == 1:
        return found[0]
    listing = ", ".join(f'{s.shape_id} "{s.name}"' for s in shapes) or "none"
    if not found:
        raise PptxSkillsError(f"no shape matches {target}; shapes on the slide: {listing}")
    ids = ", ".join(str(s.shape_id) for s in found)
    raise PptxSkillsError(f"{target} matches {len(found)} shapes (ids {ids}); target one with {{\"id\": ...}}")


def _label(changeset: dict) -> str:
    label = re.sub(r"[^a-z0-9]+", "-", str(changeset.get("label") or "").lower()).strip("-")[:40].strip("-")
    return label or "changes"


def _run(ctx: Context, ops: list) -> None:
    for number, spec in enumerate(ops, 1):
        try:
            OPS[spec["op"]](ctx, spec)
        except PptxSkillsError as exc:
            raise PptxSkillsError(f"op {number} ({spec['op']}): {exc}") from exc
        except Exception as exc:  # a bug or odd file must still leave the deck untouched and name the op
            raise PptxSkillsError(f"op {number} ({spec['op']}): unexpected {type(exc).__name__}: {exc}") from exc


def apply_changeset(deck_path: Path, changeset: dict) -> dict:
    deck_path = Path(deck_path)
    validate_changeset(changeset)
    for spec in changeset["ops"]:
        if spec["op"] not in OPS:
            raise PptxSkillsError(f"op {spec['op']} is not implemented")
    prs = open_presentation(deck_path)  # fails with a clear message if the file is missing or unreadable
    deck_dir = find_deck_dir(deck_path)
    deck = load_deck(deck_dir) if deck_dir else None
    if deck is not None and deck["working_copy"] != deck_path.name:
        deck_dir, deck = None, None  # a stray file next to deck.json, not the working copy
    messages: list[str] = []
    if deck_dir is None:
        messages.append("not inside a deck folder: no snapshot and no change-log entry were written")
    elif changed_externally(deck_dir):
        messages.append("the working copy changed outside the agent since the last apply; the change set was applied to the current file")
    if lock_file(deck_path):
        messages.append("the deck is open in PowerPoint; close it or reload it to see the result")

    try:
        before = deck_path.read_bytes()
    except OSError as exc:
        raise PptxSkillsError(f"cannot read {deck_path}: {exc}") from exc
    ctx = Context(
        prs=prs,
        deck_path=deck_path,
        deck_dir=deck_dir,
        brief=load_brief(deck_dir) if deck_dir else None,
        examples_path=deck["sources"].get("examples") if deck else None,
        messages=messages,
    )
    _run(ctx, changeset["ops"])
    live = {s.slide_id for s in ctx.prs.slides}
    touched = sorted(ctx.touched & live)
    seq = _write(ctx, changeset, before, touched)
    return {"seq": seq, "touched": touched, "messages": messages}


def _listing(folder: Path) -> set[str] | None:
    return {p.name for p in folder.iterdir()} if folder.is_dir() else None


def _write(ctx: Context, changeset: dict, before: bytes, touched: list[int]) -> int | None:
    """Replace the deck, then log it; any failure restores the deck, deck.json, changes/ and snapshots/."""
    deck_path, deck_dir = ctx.deck_path, ctx.deck_dir
    tmp = deck_path.with_name(f".{deck_path.name}.tmp")
    saved = {}
    replaced = False
    if deck_dir is not None:
        saved["json"] = (deck_dir / "deck.json").read_bytes()
        saved["dirs"] = {n: _listing(deck_dir / n) for n in ("changes", "snapshots")}
    try:
        new_hash = save_presentation(ctx.prs, tmp)
        os.replace(tmp, deck_path)
        replaced = True
        if deck_dir is None:
            return None
        return record_apply(deck_dir, changeset, _label(changeset), before, touched, new_hash)
    except OSError as exc:
        _rollback(deck_path, deck_dir, before, saved, replaced)
        raise PptxSkillsError(f"could not write {deck_path}; nothing was changed: {exc}") from exc
    except BaseException:
        _rollback(deck_path, deck_dir, before, saved, replaced)
        raise
    finally:
        tmp.unlink(missing_ok=True)


def _rollback(deck_path: Path, deck_dir: Path | None, before: bytes, saved: dict, replaced: bool) -> None:
    """Best effort: undo whatever this apply wrote."""
    steps = []
    if replaced:
        steps.append(lambda: deck_path.write_bytes(before))
    if deck_dir is not None:
        steps.append(lambda: (deck_dir / "deck.json").write_bytes(saved["json"]))
        for name, names in saved["dirs"].items():
            folder = deck_dir / name
            steps.append(lambda folder=folder, names=names: _restore_dir(folder, names))
    for step in steps:
        try:
            step()
        except OSError:
            pass


def _restore_dir(folder: Path, names: set[str] | None) -> None:
    if not folder.is_dir():
        return
    for item in folder.iterdir():
        if names is None or item.name not in names:
            item.unlink()
    if names is None:
        folder.rmdir()


@op("add_slide")
def _add_slide(ctx: Context, spec: dict) -> None:
    after = slides.resolve_slide(ctx.prs, spec["after"], ctx.last_slide) if "after" in spec else None
    slide = slides.add_slide(ctx.prs, spec["layout"], after)
    ctx.last_slide = slide
    ctx.touched.add(slide.slide_id)


@op("delete_slide")
def _delete_slide(ctx: Context, spec: dict) -> None:
    slide = slides.resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    slide_id = slide.slide_id  # unavailable once the slide is gone
    last_id = ctx.last_slide.slide_id if ctx.last_slide is not None else None
    slides.delete_slide(ctx.prs, slide)
    ctx.touched.discard(slide_id)
    if last_id == slide_id:
        ctx.last_slide = None


@op("move_slide")
def _move_slide(ctx: Context, spec: dict) -> None:
    slide = slides.resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    slides.move_slide(ctx.prs, slide, spec["to"])
    ctx.touched.add(slide.slide_id)


from . import ops_text  # noqa: E402,F401  (registers the text ops; bottom import avoids a cycle)


def _handle_apply(args) -> int:
    path = Path(args.changeset)
    try:
        changeset = json.loads(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise PptxSkillsError(f"cannot read {path}: {exc}") from exc
    except ValueError as exc:
        raise PptxSkillsError(f"{path} is not valid JSON: {exc}") from exc
    result = apply_changeset(Path(args.deck), changeset)
    lines = [f"Applied {len(changeset['ops'])} op(s)" + (f" as change {result['seq']:03d}" if result["seq"] else "")]
    if result["touched"]:
        lines.append("touched slide ids: " + ", ".join(str(i) for i in result["touched"]))
    lines += [f"note: {m}" for m in result["messages"]]
    cli.emit(args, result, "\n".join(lines))
    return cli.OK


def register(subparsers) -> None:
    apply = subparsers.add_parser("apply", help="Validate and apply a JSON change set to a working copy.")
    apply.add_argument("deck", help="working copy (.pptx)")
    apply.add_argument("changeset", help="change set (JSON file)")
    apply.add_argument("--json", action="store_true", help="machine-readable output")
    apply.set_defaults(handler=_handle_apply)
