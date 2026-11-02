"""The deck folder `decks/<slug>/`: working copy, deck.json, style brief, change log and snapshots."""

import json
import re
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from . import cli
from .brief import build_brief, load_brief, write_brief
from .errors import PptxSkillsError
from .neutral import load_style_file, neutral_presentation
from .opc import file_hash, open_presentation, save_presentation
from .slides import MIN_SLIDE_ID, delete_all_slides

_OWNED = ("changes", "snapshots", "render", "check", "style")
_SLUG = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")


def _deck_json(deck_dir: Path) -> Path:
    return Path(deck_dir) / "deck.json"


def load_deck(deck_dir: Path) -> dict:
    path = _deck_json(deck_dir)
    if not path.is_file():
        raise PptxSkillsError(f"{deck_dir} is not a deck folder (no deck.json); run `pptx init` first")
    try:
        deck = json.loads(path.read_text(encoding="utf-8"))
    except (ValueError, UnicodeDecodeError) as exc:
        raise PptxSkillsError(f"{path} is not valid JSON") from exc
    _validate(Path(deck_dir), deck, path)
    return deck


def _validate(deck_dir: Path, deck, path: Path) -> None:
    name = deck.get("working_copy") if isinstance(deck, dict) else None
    ok = (
        isinstance(name, str)
        and name.endswith(".pptx")
        and name not in (".", "..")
        and "/" not in name
        and "\\" not in name
        and (deck_dir / name).resolve().parent == deck_dir.resolve()
    )
    if not ok:
        raise PptxSkillsError(f"{path} has an invalid working_copy; it must be a file name inside the deck folder")
    if not isinstance(deck.get("sources"), dict):
        raise PptxSkillsError(f"{path} has invalid sources; expected an object")
    if not isinstance(deck.get("cloned_texts", {}), dict):
        raise PptxSkillsError(f"{path} has invalid cloned_texts; expected an object")
    max_id = deck.get("max_slide_id")
    if max_id is not None and (not isinstance(max_id, int) or isinstance(max_id, bool)):
        raise PptxSkillsError(f"{path} has an invalid max_slide_id; expected a whole number")


def _save_deck(deck_dir: Path, deck: dict) -> None:
    _deck_json(deck_dir).write_text(json.dumps(deck, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def find_deck_dir(pptx_path: Path) -> Path | None:
    parent = Path(pptx_path).resolve().parent
    return parent if _deck_json(parent).is_file() else None


def deck_using_as_source(path: Path) -> Path | None:
    """The deck folder that uses `path` as a style source (master, examples or style file), or None.

    Looks at the deck folder next to the file and at every `decks/*/` of the current folder.
    """
    target = Path(path).resolve()
    candidates = [target.parent, *(p.parent for p in Path.cwd().glob("decks/*/deck.json"))]
    for folder in dict.fromkeys(c.resolve() for c in candidates):
        if not _deck_json(folder).is_file():
            continue
        try:
            sources = load_deck(folder)["sources"].values()
        except PptxSkillsError:
            continue  # a broken deck.json names no sources we could trust
        if any(src and Path(src).resolve() == target for src in sources):
            return folder
    return None


def _abs(path: Path | None) -> str | None:
    return str(Path(path).resolve()) if path is not None else None


def _style_rules(style_file: str | None) -> dict | None:
    return load_style_file(Path(style_file)).get("rules") if style_file else None


def refresh_brief(
    deck_dir: Path,
    deck: dict,
    force: bool = False,
    notes: list[str] | None = None,
    renderer: str | None = None,
    sheets: bool = False,
) -> dict:
    """Rebuild style/brief.json and brief.md; with `sheets` also the contact sheets of layouts and examples
    (skipped while the sources are unchanged).

    `force` re-renders the sheets with `renderer` (default: PPTX_SKILLS_RENDERER, else auto) and implies `sheets`;
    messages for the user (e.g. skipped sheets) are appended to `notes`.
    """
    from .render import update_sheets

    deck_dir = Path(deck_dir)
    master = open_presentation(deck_dir / deck["working_copy"])
    examples_path = deck["sources"].get("examples")
    examples = open_presentation(examples_path) if examples_path else None
    brief = build_brief(master, examples, _style_rules(deck["sources"].get("style_file")))
    try:
        previous = (load_brief(deck_dir) or {}).get("sources_hash")
    except PptxSkillsError:
        previous = None
    if sheets or force:
        brief["sources_hash"], sheet_notes = update_sheets(deck_dir, deck, previous, force, renderer)
    else:  # sheets that exist stay; the old hash lets a later --sheets see whether they are stale
        brief["sources_hash"], sheet_notes = previous, []
    write_brief(brief, deck_dir / "style")
    if notes is not None:
        notes.extend(sheet_notes)
    return brief


def _clear(deck_dir: Path, deck: dict) -> None:
    """Remove what the old deck.json describes: its working copy, the tool's own subfolders and deck.json.

    deck.json goes last, so a failure leaves a deck folder that `init --replace` can clear again.
    """
    for path in (deck_dir / deck["working_copy"], *(deck_dir / name for name in _OWNED), _deck_json(deck_dir)):
        try:
            if path.is_dir() and not path.is_symlink():
                shutil.rmtree(path)
            elif path.exists() or path.is_symlink():
                path.unlink()
        except OSError as exc:
            raise PptxSkillsError(
                f"could not remove {path} while replacing the deck: {exc}; remove it by hand and run init with --replace again"
            ) from exc


def _slug(deck_dir: Path) -> str:
    slug = deck_dir.name
    if not _SLUG.match(slug):
        raise PptxSkillsError(
            f"the deck folder name {str(deck_dir)!r} is not a kebab-case slug; name the folder after the deck, "
            "for example decks/ai-governance (lowercase letters, digits and hyphens)"
        )
    return slug


def init_deck(
    deck_dir: Path,
    source: Path | None,
    examples: Path | None = None,
    style_file: Path | None = None,
    keep_slides: bool = False,
    mode: str = "new",
    notes: list[str] | None = None,
    renderer: str | None = None,
    sheets: bool = False,
) -> dict:
    """Create (or reuse or replace) the deck folder. Without `source` and `style_file` the working copy starts on
    the neutral 16:9 template. `sheets` renders the contact sheets with `renderer` (see refresh_brief)."""
    deck_dir = Path(deck_dir)
    exists = _deck_json(deck_dir).is_file()
    if mode == "new" and exists:
        raise PptxSkillsError(f"{deck_dir} already has a deck.json; use --reuse or --replace")
    if mode == "reuse" and exists:
        deck = load_deck(deck_dir)
        for key, value in (("master", source), ("examples", examples), ("style_file", style_file)):
            if value is not None:
                deck["sources"][key] = _abs(value)
        refresh_brief(deck_dir, deck, notes=notes, renderer=renderer, sheets=sheets)
        _save_deck(deck_dir, deck)
        return deck

    slug = _slug(deck_dir)
    working_copy = f"{slug}.pptx"
    old = None
    if mode == "replace":
        if not exists:
            raise PptxSkillsError(
                f"{deck_dir} is not a deck folder (no deck.json), so there is no deck to start over from; "
                "init the deck in a new or empty folder such as decks/<slug>"
            )
        old = load_deck(deck_dir)
        owned = {old["working_copy"], *_OWNED}
    else:
        owned = set()
    clash = [n for n in (working_copy, *_OWNED) if (deck_dir / n).exists() and n not in owned]
    if clash:
        raise PptxSkillsError(
            f"{deck_dir} already contains {', '.join(clash)}, which init would overwrite; "
            "init the deck in a new or empty folder such as decks/<slug>"
        )
    # validate every input before touching the folder
    style = load_style_file(style_file) if style_file is not None else None
    prs = open_presentation(source) if source is not None else neutral_presentation(style)
    if examples is not None:
        open_presentation(examples)
    if not keep_slides:
        delete_all_slides(prs)
    deck_dir.parent.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix=f".{slug}-init-", dir=deck_dir.parent))
    try:
        deck = {
            "slug": slug,
            "working_copy": working_copy,
            "sources": {"master": _abs(source), "examples": _abs(examples), "style_file": _abs(style_file)},
            "keep_slides": keep_slides,
            "last_hash": save_presentation(prs, stage / working_copy),
            "seq": 0,
            "last_touched": [],
            # the highest slide id ever used here; new slides get higher ids (R26)
            "max_slide_id": max([MIN_SLIDE_ID - 1, *(s.slide_id for s in prs.slides), (old or {}).get("max_slide_id") or 0]),
        }
        refresh_brief(stage, deck, notes=notes, renderer=renderer, sheets=sheets)
        _save_deck(stage, deck)
        deck_dir.mkdir(parents=True, exist_ok=True)
        if old is not None:
            _clear(deck_dir, old)
        for item in stage.iterdir():
            shutil.move(str(item), str(deck_dir / item.name))
    finally:
        shutil.rmtree(stage, ignore_errors=True)
    return deck


def record_apply(
    deck_dir: Path,
    changeset: dict | str,
    label: str,
    before: bytes,
    touched: list[int],
    new_hash: str,
    cloned_texts: dict | None = None,
    max_slide_id: int | None = None,
) -> int:
    """Log the change set (a dict, logged as `changes/NNN-<label>.json`, or a script's text, logged as `.py`), keep
    the snapshot and update deck.json.

    `cloned_texts` replaces the stored record; `max_slide_id` raises the stored highest slide id, never lowers it.
    """
    deck_dir = Path(deck_dir)
    deck = load_deck(deck_dir)
    number = deck["seq"] + 1
    slug = re.sub(r"[^a-z0-9]+", "-", label.lower()).strip("-") or "change"
    (deck_dir / "changes").mkdir(exist_ok=True)
    (deck_dir / "snapshots").mkdir(exist_ok=True)
    if isinstance(changeset, str):
        (deck_dir / "changes" / f"{number:03d}-{slug}.py").write_text(changeset, encoding="utf-8")
    else:
        (deck_dir / "changes" / f"{number:03d}-{slug}.json").write_text(
            json.dumps(changeset, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
        )
    (deck_dir / "snapshots" / f"{number:03d}.pptx").write_bytes(before)
    deck.update(
        seq=number,
        last_hash=new_hash,
        last_touched=list(touched),
        last_applied=datetime.now(timezone.utc).isoformat(timespec="seconds"),
    )
    if max_slide_id is not None:
        deck["max_slide_id"] = max(deck.get("max_slide_id") or 0, max_slide_id)
    if cloned_texts is not None:
        deck.pop("cloned_texts", None)
        if cloned_texts:
            deck["cloned_texts"] = cloned_texts
    _save_deck(deck_dir, deck)
    return number


def changed_externally(deck_dir: Path) -> bool:
    deck = load_deck(deck_dir)
    return file_hash(Path(deck_dir) / deck["working_copy"]) != deck["last_hash"]


def _summary(deck: dict, deck_dir: Path) -> str:
    text = f"Deck {deck['slug']} ready: {deck['working_copy']}, style/brief.md"
    if not deck["sources"].get("master"):
        text += (f"\nIt starts on the neutral 16:9 template (no master); set its colors and fonts with "
                 f"`pptx theme {deck_dir} --color accent1=#RRGGBB --font major=<family>`")
    return text


def _handle_init(args) -> int:
    notes: list[str] = []
    mode = "reuse" if args.reuse else "replace" if args.replace else "new"
    deck = init_deck(
        Path(args.deck_dir),
        Path(args.source) if args.source else None,
        Path(args.examples) if args.examples else None,
        Path(args.style_file) if args.style_file else None,
        args.keep_slides,
        mode,
        notes,
        args.renderer,
        args.sheets,
    )
    cli.emit(args, {**deck, "notes": notes} if notes else deck, "\n".join([_summary(deck, Path(args.deck_dir)), *notes]))
    return cli.OK


def _handle_style(args) -> int:
    deck_dir = Path(args.deck_dir)
    deck = load_deck(deck_dir)
    notes: list[str] = []
    refresh_brief(deck_dir, deck, force=args.refresh, notes=notes, renderer=args.renderer, sheets=args.sheets)
    style = deck_dir / "style"
    data = {"brief": str(style / "brief.json"), "markdown": str(style / "brief.md"), "notes": notes}
    for key, name in (("layouts", "layouts.png"), ("examples", "examples.png")):
        if (style / name).is_file():
            data[key] = str(style / name)
    cli.emit(args, data, "\n".join([f"Style brief written to {data['markdown']}", *notes]))
    return cli.OK


def register(subparsers) -> None:
    from .render import add_renderer_option

    init = subparsers.add_parser("init", help="Create a deck folder with a working copy and a style brief.")
    init.add_argument("deck_dir", help="deck folder, e.g. decks/<slug>")
    init.add_argument(
        "--from", dest="source",
        help="master or deck (.potx/.pptx) whose layouts define the style; without it the deck starts on a neutral 16:9 template",
    )
    init.add_argument("--examples", help="example deck (.pptx) to learn style usage from")
    init.add_argument("--style-file", help="style file (JSON) with check rules")
    init.add_argument("--keep-slides", action="store_true", help="keep the slides of the source deck")
    group = init.add_mutually_exclusive_group()
    group.add_argument("--reuse", action="store_true", help="keep the working copy, refresh the brief")
    group.add_argument("--replace", action="store_true", help="start over in an existing deck folder; files the deck does not own stay")
    init.add_argument("--sheets", action="store_true", help="also render contact sheets of the layouts and examples to style/")
    add_renderer_option(init)
    init.add_argument("--json", action="store_true", help="machine-readable output")
    init.set_defaults(handler=_handle_init)

    style = subparsers.add_parser("style", help="Rebuild the style brief of a deck folder.")
    style.add_argument("deck_dir", help="deck folder")
    style.add_argument("--sheets", action="store_true", help="also render contact sheets of the layouts and examples to style/")
    style.add_argument("--refresh", action="store_true", help="re-render the contact sheets even if the sources are unchanged")
    add_renderer_option(style)
    style.add_argument("--json", action="store_true", help="machine-readable output")
    style.set_defaults(handler=_handle_style)
