"""Apply a validated JSON change set to a working copy: all ops succeed or nothing is written."""

import copy
import json
import os
import re
from dataclasses import dataclass, field
from importlib import resources
from pathlib import Path

from jsonschema import Draft202012Validator
from pptx.oxml.ns import qn

from . import cli, slides
from .brief import load_brief
from .deckdir import changed_externally, deck_using_as_source, find_deck_dir, load_deck, record_apply
from .errors import PptxSkillsError
from .geometry import absolute_box
from .inventory import shape_text, table_cell_texts
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
    # deck.json `cloned_texts` as it will be written: {slide id: {shape id: example text}}
    cloned_texts: dict = field(default_factory=dict)
    sources: dict = field(default_factory=dict)  # resolved path -> opened source deck
    # the highest slide id the deck folder has used, deleted slides included; new slides go above it (R26)
    max_slide_id: int = 0


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
        if shape._element.tag == qn("p:grpSp"):
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


def _check_target(deck_path: Path, deck: dict | None, outside_deck: bool) -> None:
    """Refuse to write anything but a working copy (R28): never a template or a deck's style source, and any
    other file only with `outside_deck`. `deck` is the deck.json of the folder whose working copy this is."""
    if deck_path.suffix.lower() == ".potx":
        raise PptxSkillsError(
            f"{deck_path} is a template, and apply never writes a .potx; start a deck from it with "
            f"`pptx init decks/<slug> --from {deck_path}` and apply the change set to its working copy"
        )
    target = deck_path.resolve()
    if deck is not None:  # a working copy: only its own deck's sources count (another deck may start from it)
        if any(src and Path(src).resolve() == target for src in deck["sources"].values()):
            raise PptxSkillsError(f"{deck_path} is also a style source of its own deck folder; apply never writes a source")
        return
    owner = deck_using_as_source(deck_path)
    if owner is not None:
        raise PptxSkillsError(
            f"{deck_path} is a style source of the deck folder {owner}, and apply never writes the user's original; "
            f"apply the change set to that deck's working copy instead"
        )
    if not outside_deck:
        raise PptxSkillsError(
            f"{deck_path} is not the working copy of a deck folder, and apply writes only working copies so the "
            f"user's own file stays untouched; start a deck from it with `pptx init decks/<slug> --from {deck_path}` "
            "(add --keep-slides to keep its slides), or pass --outside-deck to change this file in place without "
            "snapshot or change log"
        )


def apply_changeset(deck_path: Path, changeset: dict, outside_deck: bool = False) -> dict:
    """Apply `changeset` to the working copy `deck_path`; `outside_deck` also allows any other .pptx (R28)."""
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
    _check_target(deck_path, deck, outside_deck)
    messages: list[str] = []
    if deck_dir is None:
        messages.append("--outside-deck: the file was changed in place; no snapshot and no change-log entry are kept")
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
        cloned_texts=copy.deepcopy(deck.get("cloned_texts") or {}) if deck else {},
        max_slide_id=max([(deck or {}).get("max_slide_id") or 0, *(s.slide_id for s in prs.slides)]),
    )
    _run(ctx, changeset["ops"])
    live = {s.slide_id for s in ctx.prs.slides}
    touched = sorted(ctx.touched & live)
    live_keys = {str(i) for i in live}
    ctx.cloned_texts = {k: v for k, v in ctx.cloned_texts.items() if k in live_keys}  # also drops deleted slides
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
        return record_apply(deck_dir, changeset, _label(changeset), before, touched, new_hash, ctx.cloned_texts, ctx.max_slide_id)
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
    slide = slides.add_slide(ctx.prs, spec["layout"], after, ctx.max_slide_id, ctx.messages)
    ctx.max_slide_id = max(ctx.max_slide_id, slide.slide_id)
    ctx.last_slide = slide
    ctx.touched.add(slide.slide_id)
    # ids are new (R26), except past the largest valid id, where one may be reused: never inherit example texts
    ctx.cloned_texts.pop(str(slide.slide_id), None)


@op("delete_slide")
def _delete_slide(ctx: Context, spec: dict) -> None:
    slide = slides.resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    slide_id = slide.slide_id  # unavailable once the slide is gone
    last_id = ctx.last_slide.slide_id if ctx.last_slide is not None else None
    slides.delete_slide(ctx.prs, slide, ctx.messages)
    ctx.touched.discard(slide_id)
    if last_id == slide_id:
        ctx.last_slide = None


@op("move_slide")
def _move_slide(ctx: Context, spec: dict) -> None:
    slide = slides.resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    slides.move_slide(ctx.prs, slide, spec["to"])
    ctx.touched.add(slide.slide_id)


def _source(ctx: Context, ref: dict):
    """(presentation, slide, kind) for a `from` ref; kind is "examples", "self" or "other"."""
    deck = ref["deck"]
    examples = Path(ctx.examples_path).resolve() if ctx.examples_path else None
    if deck == "self":
        prs, kind = ctx.prs, "self"
    else:
        if deck == "examples":
            if ctx.deck_dir is None:
                raise PptxSkillsError(f'"examples" needs the examples deck of a deck folder; {ctx.deck_path.name} is not inside a deck folder, so give the path of the deck instead')
            if examples is None:
                raise PptxSkillsError(f"this deck folder has no examples deck; add one with `pptx init {ctx.deck_dir} --reuse --examples <deck.pptx>` or give the path of the deck")
            path = examples
        else:
            path = Path(deck).resolve()
        if path == ctx.deck_path.resolve():
            prs, kind = ctx.prs, "self"
        else:
            if str(path) not in ctx.sources:
                ctx.sources[str(path)] = open_presentation(path)
            prs, kind = ctx.sources[str(path)], "examples" if path == examples else "other"
    try:
        slide = slides.resolve_slide(prs, ref["slide"], ctx.last_slide if kind == "self" else None)
    except PptxSkillsError as exc:
        raise PptxSkillsError(f"from {deck}: {exc}") from exc
    return prs, slide, kind


def _example_texts(slide) -> dict[str, str]:
    """{shape id: text} of a slide's text shapes, without the master's date, footer and slide number.

    Table cells are recorded as `"<shape id>:r<row>c<col>"` (0-based).
    """
    texts = {}
    for shape in _walk(slide.shapes):
        if getattr(shape, "has_table", False) and shape.has_table:
            for r, row in enumerate(table_cell_texts(shape)):
                texts.update({f"{shape.shape_id}:r{r}c{c}": t for c, t in enumerate(row) if t.strip()})
            continue
        if shape.is_placeholder and shape.placeholder_format.type in slides.FOOTER_TYPES:
            continue
        text = shape_text(shape)
        if text.strip():
            texts[str(shape.shape_id)] = text
    return texts


def _copied(ctx: Context, slide, kind: str, carried: dict | None, deck: str, notes: list[str]) -> None:
    """Bookkeeping after a slide was copied in: `$last`, touched ids, notes and the recorded example texts.

    Text from the examples deck is recorded so `check` flags it while it is unchanged; a duplicate keeps
    the record of its source (`carried`); slides from the user's own decks carry real content and record nothing.
    """
    ctx.last_slide = slide
    ctx.touched.add(slide.slide_id)
    ctx.messages.extend(f"from {deck}: {note}" for note in notes)
    record = _example_texts(slide) if kind == "examples" else carried
    if record:
        ctx.cloned_texts[str(slide.slide_id)] = copy.deepcopy(record)
    else:
        ctx.cloned_texts.pop(str(slide.slide_id), None)


def _carried_texts(ctx: Context, slide, kind: str) -> dict | None:
    """The example-text record a duplicate of `slide` inherits."""
    return ctx.cloned_texts.get(str(slide.slide_id)) if kind == "self" else None


@op("clone_slide")
def _clone_slide(ctx: Context, spec: dict) -> None:
    src_prs, src_slide, kind = _source(ctx, spec["from"])
    after = slides.resolve_slide(ctx.prs, spec["after"], ctx.last_slide) if "after" in spec else None
    carried = _carried_texts(ctx, src_slide, kind)
    notes: list[str] = []
    slide = slides.clone_slide(src_prs, src_slide, ctx.prs, after, notes, ctx.max_slide_id)
    ctx.max_slide_id = max(ctx.max_slide_id, slide.slide_id)
    _copied(ctx, slide, kind, carried, spec["from"]["deck"], notes)


@op("replace_slide")
def _replace_slide(ctx: Context, spec: dict) -> None:
    target = slides.resolve_slide(ctx.prs, spec["slide"], ctx.last_slide)
    src_prs, src_slide, kind = _source(ctx, spec["from"])
    carried = _carried_texts(ctx, src_slide, kind)  # read now: replacing a slide by itself drops the old one
    notes: list[str] = []
    slide = slides.replace_slide(ctx.prs, target, src_prs, src_slide, notes)
    _copied(ctx, slide, kind, carried, spec["from"]["deck"], notes)


# Ops modules register themselves on import; the bottom import avoids a cycle. Modules: ops_text, ops_shapes, ops_tables, ops_media.
from . import ops_media, ops_shapes, ops_tables, ops_text  # noqa: E402,F401


def _handle_apply(args) -> int:
    path = Path(args.changeset)
    try:
        changeset = json.loads(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise PptxSkillsError(f"cannot read {path}: {exc}") from exc
    except ValueError as exc:
        raise PptxSkillsError(f"{path} is not valid JSON: {exc}") from exc
    result = apply_changeset(Path(args.deck), changeset, args.outside_deck)
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
    apply.add_argument(
        "--outside-deck", action="store_true",
        help="also change a .pptx that is not a deck folder's working copy, in place and without snapshot or change log "
        "(never a .potx or a deck's style source)",
    )
    apply.add_argument("--json", action="store_true", help="machine-readable output")
    apply.set_defaults(handler=_handle_apply)
