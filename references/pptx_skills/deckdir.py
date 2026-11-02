"""The deck folder `decks/<slug>/`: working copy, deck.json, style brief, change log and snapshots."""

import json
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path

from . import cli
from .brief import build_brief, write_brief
from .errors import PptxSkillsError
from .opc import file_hash, open_presentation, save_presentation
from .slides import delete_slide

_OWNED = ("changes", "snapshots", "render", "check", "style")


def _deck_json(deck_dir: Path) -> Path:
    return Path(deck_dir) / "deck.json"


def load_deck(deck_dir: Path) -> dict:
    path = _deck_json(deck_dir)
    if not path.is_file():
        raise PptxSkillsError(f"{deck_dir} is not a deck folder (no deck.json); run `pptx init` first")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (ValueError, UnicodeDecodeError) as exc:
        raise PptxSkillsError(f"{path} is not valid JSON") from exc


def _save_deck(deck_dir: Path, deck: dict) -> None:
    _deck_json(deck_dir).write_text(json.dumps(deck, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def find_deck_dir(pptx_path: Path) -> Path | None:
    parent = Path(pptx_path).resolve().parent
    return parent if _deck_json(parent).is_file() else None


def _abs(path: Path | None) -> str | None:
    return str(Path(path).resolve()) if path is not None else None


def _style_rules(style_file: str | None) -> dict | None:
    if not style_file:
        return None
    path = Path(style_file)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise PptxSkillsError(f"cannot read style file {path}: {exc}") from exc
    except ValueError as exc:
        raise PptxSkillsError(f"style file {path} is not valid JSON: {exc}") from exc
    rules = data.get("rules") if isinstance(data, dict) else None
    return rules if isinstance(rules, dict) else None


def refresh_brief(deck_dir: Path, deck: dict) -> dict:
    """Rebuild style/brief.json and brief.md from the working copy, the examples and the style file."""
    deck_dir = Path(deck_dir)
    master = open_presentation(deck_dir / deck["working_copy"])
    examples_path = deck["sources"].get("examples")
    examples = open_presentation(examples_path) if examples_path else None
    brief = build_brief(master, examples, _style_rules(deck["sources"].get("style_file")))
    write_brief(brief, deck_dir / "style")
    return brief


def _clear(deck_dir: Path, deck: dict | None) -> None:
    if deck is not None:
        (deck_dir / deck["working_copy"]).unlink(missing_ok=True)
    _deck_json(deck_dir).unlink(missing_ok=True)
    for name in _OWNED:
        shutil.rmtree(deck_dir / name, ignore_errors=True)


def init_deck(
    deck_dir: Path,
    source: Path | None,
    examples: Path | None = None,
    style_file: Path | None = None,
    keep_slides: bool = False,
    mode: str = "new",
) -> dict:
    deck_dir = Path(deck_dir)
    exists = _deck_json(deck_dir).is_file()
    if mode == "new" and exists:
        raise PptxSkillsError(f"{deck_dir} already has a deck.json; use --reuse or --replace")
    if mode == "reuse" and exists:
        deck = load_deck(deck_dir)
        for key, value in (("master", source), ("examples", examples), ("style_file", style_file)):
            if value is not None:
                deck["sources"][key] = _abs(value)
        refresh_brief(deck_dir, deck)
        _save_deck(deck_dir, deck)
        return deck
    if source is None and style_file is None:
        raise PptxSkillsError("init needs --from <master.potx|pptx> or --style-file <style.json>")
    if source is None:
        raise PptxSkillsError("decks from a style file alone are not supported yet")

    slug = deck_dir.name
    working_copy = deck_dir / f"{slug}.pptx"
    prs = open_presentation(source)  # fail before touching the folder
    if examples is not None:
        open_presentation(examples)
    if mode == "replace" and exists:
        _clear(deck_dir, load_deck(deck_dir))
    deck_dir.mkdir(parents=True, exist_ok=True)
    if not keep_slides:
        for slide in list(prs.slides):
            delete_slide(prs, slide)
    last_hash = save_presentation(prs, working_copy)
    deck = {
        "slug": slug,
        "working_copy": working_copy.name,
        "sources": {"master": _abs(source), "examples": _abs(examples), "style_file": _abs(style_file)},
        "keep_slides": keep_slides,
        "last_hash": last_hash,
        "seq": 0,
        "last_touched": [],
    }
    refresh_brief(deck_dir, deck)
    _save_deck(deck_dir, deck)
    return deck


def record_apply(deck_dir: Path, changeset: dict, label: str, before: bytes, touched: list[int], new_hash: str) -> int:
    deck_dir = Path(deck_dir)
    deck = load_deck(deck_dir)
    number = deck["seq"] + 1
    slug = re.sub(r"[^a-z0-9]+", "-", label.lower()).strip("-") or "change"
    (deck_dir / "changes").mkdir(exist_ok=True)
    (deck_dir / "snapshots").mkdir(exist_ok=True)
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
    _save_deck(deck_dir, deck)
    return number


def changed_externally(deck_dir: Path) -> bool:
    deck = load_deck(deck_dir)
    return file_hash(Path(deck_dir) / deck["working_copy"]) != deck["last_hash"]


def _summary(deck: dict) -> str:
    return f"Deck {deck['slug']} ready: {deck['working_copy']}, style/brief.md"


def _handle_init(args) -> int:
    mode = "reuse" if args.reuse else "replace" if args.replace else "new"
    deck = init_deck(
        Path(args.deck_dir),
        Path(args.source) if args.source else None,
        Path(args.examples) if args.examples else None,
        Path(args.style_file) if args.style_file else None,
        args.keep_slides,
        mode,
    )
    cli.emit(args, deck, _summary(deck))
    return cli.OK


def _handle_style(args) -> int:
    deck_dir = Path(args.deck_dir)
    deck = load_deck(deck_dir)
    refresh_brief(deck_dir, deck)  # --refresh gains meaning with the contact sheets (Task 12)
    data = {"brief": str(deck_dir / "style" / "brief.json"), "markdown": str(deck_dir / "style" / "brief.md")}
    cli.emit(args, data, f"Style brief written to {data['markdown']}")
    return cli.OK


def register(subparsers) -> None:
    init = subparsers.add_parser("init", help="Create a deck folder with a working copy and a style brief.")
    init.add_argument("deck_dir", help="deck folder, e.g. decks/<slug>")
    init.add_argument("--from", dest="source", help="master or deck (.potx/.pptx) whose layouts define the style")
    init.add_argument("--examples", help="example deck (.pptx) to learn style usage from")
    init.add_argument("--style-file", help="style file (JSON) with check rules")
    init.add_argument("--keep-slides", action="store_true", help="keep the slides of the source deck")
    group = init.add_mutually_exclusive_group()
    group.add_argument("--reuse", action="store_true", help="keep the working copy, refresh the brief")
    group.add_argument("--replace", action="store_true", help="start over inside the deck folder")
    init.add_argument("--json", action="store_true", help="machine-readable output")
    init.set_defaults(handler=_handle_init)

    style = subparsers.add_parser("style", help="Rebuild the style brief of a deck folder.")
    style.add_argument("deck_dir", help="deck folder")
    style.add_argument("--refresh", action="store_true", help="rebuild cached style artifacts")
    style.add_argument("--json", action="store_true", help="machine-readable output")
    style.set_defaults(handler=_handle_style)
