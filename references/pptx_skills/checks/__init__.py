"""`pptx check`: deterministic rules an agent runs after every change, before it shows a slide.

Rules register with `@rule(rule_id, category, severity)` and take `(ctx, slide)`; the modules
`structure`, `text`, `geometry` and `style` are imported at the bottom of this file. By default only
DEFAULT_RULE_IDS run; `strict` runs every rule.
"""

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path

from pptx.oxml.ns import qn

from .. import cli
from ..brief import DEFAULT_RULES, load_brief
from ..deckdir import find_deck_dir, load_deck
from ..errors import PptxSkillsError
from ..fonts import _Inheritance, measure_paragraphs
from ..geometry import Box, absolute_box, visual_box
from ..inventory import inventory
from ..opc import open_presentation

CATEGORIES = ("structure", "text", "geometry", "style")
SEVERITIES = ("error", "warning", "info")
# What every check runs: errors that break a slide or the file, plus two writing rules the user insists on
# (no keyword chains, no walls of text). Matters of taste (alignment, sizes, colors, fonts) run with `strict`.
DEFAULT_RULE_IDS = frozenset({
    "reopen", "relationships", "duplicate-ids", "overflow", "off-slide", "leftover-text",
    "keyword-chain", "too-many-lines",
})


@dataclass
class Finding:
    rule: str
    category: str
    severity: str
    slide: int | None
    slide_id: int | None
    shape_id: int | None
    shape: str | None
    message: str
    fix: str


@dataclass
class Rule:
    rule_id: str
    category: str
    severity: str
    func: object
    scope: str = "slide"


RULES: list[Rule] = []


def rule(rule_id: str, category: str, severity: str, scope: str = "slide"):
    """Register a rule; scope "deck" rules run once with `slide=None`."""
    assert category in CATEGORIES and severity in SEVERITIES and scope in ("slide", "deck")

    def decorate(func):
        RULES.append(Rule(rule_id, category, severity, func, scope))
        return func

    return decorate


# measured text may exceed its box by this much before it counts (rounding, font metrics)
TEXT_FIT_TOLERANCE_CM = 0.05


@dataclass
class SlideView:
    """One slide under check: its python-pptx slide, number, id and inventory entry."""

    number: int
    id: int
    pptx: object
    info: dict
    _grown: dict = field(default_factory=dict, repr=False)

    def grown_height(self, shape) -> float | None:
        """The height in cm of a shape that grows to fit its text (`spAutoFit`), when it exceeds the stored height."""
        if shape.shape_id not in self._grown:
            self._grown[shape.shape_id] = self._measure_growth(shape)
        return self._grown[shape.shape_id]

    def _measure_growth(self, shape) -> float | None:
        tx_body = shape._element.find(qn("p:txBody"))
        if tx_body is None or not "".join(t.text or "" for t in tx_body.iter(qn("a:t"))).strip():
            return None
        if _Inheritance(shape, self.pptx).body().autofit != "shape":
            return None
        height = measure_paragraphs(shape, self.pptx).height_cm
        return height if height > absolute_box(shape).h + TEXT_FIT_TOLERANCE_CM else None

    def seen_box(self, shape) -> Box:
        """What the viewer sees: the rotation-aware bounds (R21) at the grown height of a growing text box."""
        return visual_box(shape, self.grown_height(shape))

    def shapes(self):
        """(inventory shape dict, python-pptx shape) for every shape, groups flattened."""
        by_id = {}

        def walk(shapes):
            for shape in shapes:
                by_id[shape.shape_id] = shape
                if shape._element.tag == qn("p:grpSp"):
                    walk(shape.shapes)

        walk(self.pptx.shapes)

        def flat(items):
            for item in items:
                yield item
                yield from flat(item["children"])

        for item in flat(self.info["shapes"]):
            if item["id"] in by_id:
                yield item, by_id[item["id"]]


@dataclass
class Context:
    prs: object
    path: Path
    rules: dict
    content_area: list[float] | None
    deck: dict = field(default_factory=dict)
    _current: Rule | None = None

    def finding(self, slide: SlideView | None, shape: dict | None, message: str, fix: str) -> Finding:
        rule_ = self._current
        return Finding(
            rule=rule_.rule_id,
            category=rule_.category,
            severity=rule_.severity,
            slide=slide.number if slide else None,
            slide_id=slide.id if slide else None,
            shape_id=shape["id"] if shape else None,
            shape=shape["name"] if shape else None,
            message=message,
            fix=fix,
        )


def _resolve(deck_path: Path) -> tuple[Path, Path | None]:
    """(pptx file, deck folder or None) for a .pptx file or a deck folder."""
    deck_path = Path(deck_path)
    if deck_path.is_dir():
        deck = load_deck(deck_path)
        return deck_path / deck["working_copy"], deck_path
    return deck_path, find_deck_dir(deck_path)


def _selected(prs, deck: dict, slides, all_slides: bool) -> list[int]:
    """The slide numbers to check: `slides` as given ([] = none, deck rules only), else the last touched or all."""
    count = len(prs.slides)
    if slides is not None:
        bad = [n for n in slides if not 1 <= n <= count]
        if bad:
            raise PptxSkillsError(f"slide {bad[0]} does not exist; the deck has {count} slides")
        return sorted(set(slides))
    if not all_slides:
        ids = {s.slide_id: n for n, s in enumerate(prs.slides, 1)}
        touched = sorted(ids[i] for i in deck.get("last_touched", []) if i in ids)
        if touched:
            return touched
    return list(range(1, count + 1))


def run_checks(deck_path: Path, slides: list[int] | None = None, all_slides: bool = False, strict: bool = False) -> dict:
    """Check the slides (see `_selected`) with DEFAULT_RULE_IDS, or with every rule when `strict`.

    Without `strict` a theme font that is not installed is a note in `report["notes"]`, not a finding.
    """
    path, deck_dir = _resolve(deck_path)
    prs = open_presentation(path)
    deck = load_deck(deck_dir) if deck_dir else {}
    brief = load_brief(deck_dir) if deck_dir else None
    rules = {**DEFAULT_RULES, **((brief or {}).get("rules") or {})}
    ctx = Context(prs, path, rules, (brief or {}).get("content_area") or None, deck)
    numbers = _selected(prs, deck, slides, all_slides)
    inv = inventory(prs, numbers)
    views = [SlideView(s["number"], s["id"], prs.slides[s["number"] - 1], s) for s in inv["slides"]]

    findings: list[Finding] = []
    for item in RULES:
        if not strict and item.rule_id not in DEFAULT_RULE_IDS:
            continue
        ctx._current = item
        if item.scope == "deck":
            findings.extend(item.func(ctx, None))
        else:
            for view in views:
                findings.extend(item.func(ctx, view))
    summary = {s: sum(f.severity == s for f in findings) for s in SEVERITIES}
    report = {
        "deck": str(path),
        "slides": numbers,
        "findings": [asdict(f) for f in sorted(findings, key=lambda f: (f.slide or 0, SEVERITIES.index(f.severity)))],
        "summary": summary,
    }
    missing = [] if strict else style.missing_theme_fonts(prs)  # with strict: the finding missing-font
    if missing:
        report["notes"] = [
            f"the theme font {' and '.join(missing)} is not installed, so text is measured and rendered with a "
            "fallback font; line breaks may differ in PowerPoint"
        ]
    if deck_dir:
        _write_slide_reports(deck_dir, views, report["findings"])
    return report


def _write_slide_reports(deck_dir: Path, views: list[SlideView], findings: list[dict]) -> None:
    out = Path(deck_dir) / "check"
    out.mkdir(exist_ok=True)
    for view in views:
        data = {"slide": view.number, "slide_id": view.id, "findings": [f for f in findings if f["slide"] == view.number]}
        (out / f"slide-{view.number:02d}.json").write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def format_report(report: dict) -> str:
    lines = []
    for f in report["findings"]:
        where = f"slide {f['slide']}" if f["slide"] else "deck"
        if f["shape"]:
            where += f" [{f['shape']}]"
        lines.append(f"{where} {f['severity']} {f['rule']}: {f['message']} — {f['fix']}")
    lines += [f"note: {note}" for note in report.get("notes", [])]
    s = report["summary"]
    shown = ",".join(str(n) for n in report["slides"]) or "none"
    lines.append(f"Checked slides {shown}: {s['error']} errors, {s['warning']} warnings, {s['info']} info.")
    return "\n".join(lines)


def _handle(args) -> int:
    slides = None
    if args.slides:
        try:
            slides = [int(n) for n in args.slides.split(",")]
        except ValueError:
            raise PptxSkillsError(f"--slides expects numbers like 1,3, got {args.slides!r}") from None
    report = run_checks(Path(args.deck), slides, args.all, args.strict)
    cli.emit(args, report, format_report(report))
    return cli.FINDINGS if report["summary"]["error"] else cli.OK


def register(subparsers) -> None:
    parser = subparsers.add_parser(
        "check", help="Run the slide checks: real errors plus keyword chains and long paragraphs; exit 1 on errors."
    )
    parser.add_argument("deck", help="deck folder or .pptx file")
    parser.add_argument("--slides", help="comma-separated slide numbers (default: slides touched last, else all)")
    parser.add_argument("--all", action="store_true", help="check every slide")
    parser.add_argument(
        "--strict", action="store_true",
        help="run every rule: also sizes, alignment, gaps, overlaps, colors, fonts, alt text, bullets, titles",
    )
    parser.add_argument("--json", action="store_true", help="machine-readable output")
    parser.set_defaults(handler=_handle)


from . import geometry, structure, style, text  # noqa: E402,F401  (register the rules)
