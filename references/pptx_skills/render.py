"""Render slides to PNG (LibreOffice + poppler), contact sheets of layouts and examples, and `doctor`."""

import hashlib
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from . import cli
from .brief import _layouts, load_brief
from .errors import PptxSkillsError
from .fonts import find_font
from .opc import open_presentation, save_presentation
from .slides import add_slide, delete_all_slides

_SOFFICE_PATHS = (
    Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    Path.home() / "Applications/LibreOffice.app/Contents/MacOS/soffice",
)
_TIMEOUT_S = 180
# LibreOffice leaves hidden slides out of a PDF by default, so page n would no longer be slide n
_PDF_FILTER = 'pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}'
_GAP = 16
_LABEL_H = 34
_INSTALL_HINT = (
    "install LibreOffice and poppler: `brew install --cask libreoffice` (add `--appdir=~/Applications` without "
    "admin rights) and `brew install poppler`, or `apt install libreoffice poppler-utils`"
)
LAYOUTS_SHEET = "layouts.png"
EXAMPLES_SHEET = "examples.png"


def find_soffice() -> Path | None:
    found = shutil.which("soffice")
    if found:
        return Path(found)
    return next((p for p in _SOFFICE_PATHS if p.is_file()), None)


def find_pdftoppm() -> Path | None:
    found = shutil.which("pdftoppm")
    return Path(found) if found else None


def _require_tools() -> tuple[Path, Path]:
    soffice, pdftoppm = find_soffice(), find_pdftoppm()
    missing = [name for name, path in (("LibreOffice (soffice)", soffice), ("poppler (pdftoppm)", pdftoppm)) if path is None]
    if missing:
        raise PptxSkillsError(f"{' and '.join(missing)} not found; {_INSTALL_HINT}; `pptx doctor` checks the setup")
    return soffice, pdftoppm


def _run(command: list[str], what: str) -> None:
    try:
        result = subprocess.run(command, capture_output=True, text=True, errors="replace", timeout=_TIMEOUT_S)
    except subprocess.TimeoutExpired as exc:
        raise PptxSkillsError(f"{what} timed out after {_TIMEOUT_S} s") from exc
    except OSError as exc:
        raise PptxSkillsError(f"{what} could not start: {exc}") from exc
    if result.returncode != 0:
        tail = (result.stderr or result.stdout or "").strip()[-400:]
        raise PptxSkillsError(f"{what} failed (exit {result.returncode}): {tail}")


def default_out_dir(deck: Path) -> Path:
    from .deckdir import find_deck_dir, load_deck

    deck = Path(deck)
    deck_dir = find_deck_dir(deck)
    if deck_dir is not None and load_deck(deck_dir)["working_copy"] == deck.name:
        return deck_dir / "render"
    return deck.parent / f"{deck.stem}-render"


def render(deck: Path, out_dir: Path | None = None, slides: list[int] | None = None, dpi: int = 120) -> list[Path]:
    """Render the given slides (1-based, default all) to `slide-NN.png` in `out_dir`; returns the files."""
    deck = Path(deck)
    count = len(open_presentation(deck).slides)  # also validates the file
    numbers = list(slides) if slides else list(range(1, count + 1))
    bad = [n for n in numbers if not 1 <= n <= count]
    if bad:
        raise PptxSkillsError(f"slide {bad[0]} does not exist; the deck has {count} slide(s)")
    soffice, pdftoppm = _require_tools()
    out_dir = Path(out_dir) if out_dir is not None else default_out_dir(deck)
    with tempfile.TemporaryDirectory(prefix="pptx-render-") as tmp_name:
        tmp = Path(tmp_name)
        profile = tmp / "lo-profile"
        _run(
            [str(soffice), "--headless", f"-env:UserInstallation={profile.as_uri()}", "--convert-to", _PDF_FILTER,
             "--outdir", str(tmp), str(deck)],
            "LibreOffice conversion",
        )
        pdf = tmp / f"{deck.stem}.pdf"
        if not pdf.is_file():
            raise PptxSkillsError(f"LibreOffice produced no PDF for {deck}")
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise PptxSkillsError(f"cannot create output folder {out_dir}: {exc}") from exc
        written = []
        for n in numbers:
            prefix = tmp / f"page{n}"
            _run([str(pdftoppm), "-png", "-r", str(dpi), "-f", str(n), "-l", str(n), str(pdf), str(prefix)],
                 f"pdftoppm for slide {n}")
            produced = sorted(tmp.glob(f"page{n}-*.png")) + sorted(tmp.glob(f"page{n}.png"))
            if not produced:
                raise PptxSkillsError(f"pdftoppm produced no image for slide {n}")
            target = out_dir / f"slide-{n:02d}.png"
            try:
                shutil.copyfile(produced[0], target)
            except OSError as exc:
                raise PptxSkillsError(f"cannot write {target}: {exc}") from exc
            written.append(target)
    return written


def contact_sheet(pngs: list[Path], labels: list[str], out: Path, columns: int = 4, thumb_width: int = 480) -> Path:
    """Tile the images (scaled to `thumb_width`) in a grid with a caption under each."""
    if not pngs or len(pngs) != len(labels):
        raise PptxSkillsError("a contact sheet needs one label per image and at least one image")
    columns = max(1, min(columns, len(pngs)))
    thumbs = []
    for path in pngs:
        try:
            with Image.open(path) as image:
                image = image.convert("RGB")
                height = max(1, round(image.height * thumb_width / image.width))
                thumbs.append(image.resize((thumb_width, height), Image.LANCZOS))
        except (OSError, ValueError) as exc:  # UnidentifiedImageError is an OSError
            raise PptxSkillsError(f"cannot read image {path}: {exc}") from exc
    rows = [thumbs[i:i + columns] for i in range(0, len(thumbs), columns)]
    row_heights = [max(t.height for t in row) + _LABEL_H for row in rows]
    sheet = Image.new("RGB", (columns * thumb_width + (columns + 1) * _GAP, sum(row_heights) + (len(rows) + 1) * _GAP), "white")
    draw = ImageDraw.Draw(sheet)
    font = ImageFont.load_default(size=20)
    y = _GAP
    for r, row in enumerate(rows):
        for c, thumb in enumerate(row):
            x = _GAP + c * (thumb_width + _GAP)
            sheet.paste(thumb, (x, y))
            draw.rectangle([x - 1, y - 1, x + thumb.width, y + thumb.height], outline=(190, 190, 190))
            draw.text((x, y + thumb.height + 6), labels[r * columns + c], fill=(30, 30, 30), font=font)
        y += row_heights[r] + _GAP
    out = Path(out)
    try:
        out.parent.mkdir(parents=True, exist_ok=True)
        sheet.save(out)
    except (OSError, ValueError) as exc:
        raise PptxSkillsError(f"cannot write contact sheet {out}: {exc}") from exc
    return out


def render_slide_sheet(pngs: list[Path], slide_numbers: list[int], out: Path) -> Path:
    return contact_sheet(pngs, [f"Slide {n}" for n in slide_numbers], out)


def render_layouts(deck: Path, out: Path) -> Path:
    """One slide per layout of `deck` (placeholders filled with their role name), labeled with the layout name."""
    prs = open_presentation(deck)
    delete_all_slides(prs)
    layouts = _layouts(prs)
    names = [layout["name"] for layout in layouts]
    for info in layouts:
        slide = add_slide(prs, info["name"])
        roles = {ph["idx"]: ph["role"] for ph in info["placeholders"]}
        for ph in slide.placeholders:
            role = roles.get(ph.placeholder_format.idx, "other")
            if role != "picture" and ph.has_text_frame:
                ph.text_frame.text = role
    with tempfile.TemporaryDirectory(prefix="pptx-layouts-") as tmp_name:
        tmp = Path(tmp_name)
        save_presentation(prs, tmp / "layouts.pptx")
        pngs = render(tmp / "layouts.pptx", tmp / "png", dpi=60)
        return contact_sheet(pngs, names, out)


def render_examples(examples: Path, out: Path) -> Path:
    with tempfile.TemporaryDirectory(prefix="pptx-examples-") as tmp_name:
        pngs = render(examples, Path(tmp_name), dpi=60)
        if not pngs:
            raise PptxSkillsError(f"{examples} has no slides to render")
        return contact_sheet(pngs, [f"Slide {i}" for i in range(1, len(pngs) + 1)], out)


def sources_hash(deck_dir: Path, deck: dict) -> str:
    digest = hashlib.sha256()
    paths = [Path(deck_dir) / deck["working_copy"], *(Path(p) for p in (deck["sources"].get("examples"), deck["sources"].get("style_file")) if p)]
    for path in paths:
        digest.update(path.read_bytes() if path.is_file() else b"-")
        digest.update(b"\0")
    return digest.hexdigest()


def update_sheets(deck_dir: Path, deck: dict, previous_hash: str | None, force: bool = False) -> tuple[str | None, list[str]]:
    """Write style/layouts.png (and examples.png); returns (sources_hash to store, notes for the user)."""
    deck_dir = Path(deck_dir)
    style = deck_dir / "style"
    current = sources_hash(deck_dir, deck)
    examples = deck["sources"].get("examples")
    wanted = [style / LAYOUTS_SHEET] + ([style / EXAMPLES_SHEET] if examples else [])
    if not force and previous_hash == current and all(p.is_file() for p in wanted):
        return current, []
    if find_soffice() is None:
        return None, ["contact sheets skipped: LibreOffice not found — run `pptx doctor`"]
    if find_pdftoppm() is None:
        return None, ["contact sheets skipped: poppler (pdftoppm) not found — run `pptx doctor`"]
    try:
        render_layouts(deck_dir / deck["working_copy"], style / LAYOUTS_SHEET)
        if examples:
            render_examples(Path(examples), style / EXAMPLES_SHEET)
    except (PptxSkillsError, OSError) as exc:
        return None, [f"contact sheets skipped: {exc}"]
    return current, []


def doctor(deck_dir: Path | None = None) -> dict:
    fonts: dict = {}
    brief = load_brief(deck_dir) if deck_dir is not None else None
    for family in dict.fromkeys(((brief or {}).get("theme") or {}).get("fonts", {}).values()):
        if family:
            path = find_font(family)
            fonts[family] = str(path) if path else None
    soffice, pdftoppm = find_soffice(), find_pdftoppm()
    return {
        "soffice": str(soffice) if soffice else None,
        "pdftoppm": str(pdftoppm) if pdftoppm else None,
        "fonts": fonts,
    }


def _doctor_lines(report: dict) -> list[str]:
    hints = {"soffice": "brew install --cask libreoffice (or apt install libreoffice)",
             "pdftoppm": "brew install poppler (or apt install poppler-utils)"}
    lines = [f"✓ {n}: {report[n]}" if report[n] else f"✗ {n}: missing; {hints[n]}" for n in ("soffice", "pdftoppm")]
    for family, path in report["fonts"].items():
        lines.append(f"✓ font {family}: {path}" if path else f"✗ font {family}: not installed; install it or text is measured with a fallback")
    return lines


def _handle_render(args) -> int:
    deck = Path(args.deck)
    slides = None
    if args.slides:
        try:
            slides = [int(s) for s in args.slides.split(",") if s.strip()]
        except ValueError:
            raise PptxSkillsError(f"--slides expects numbers like 1,3; got {args.slides!r}") from None
    out_dir = Path(args.out) if args.out else default_out_dir(deck)
    files = render(deck, out_dir, slides, args.dpi)
    data = {"files": [str(f) for f in files]}
    if args.sheet:
        numbers = slides or list(range(1, len(files) + 1))
        data["sheet"] = str(render_slide_sheet(files, numbers, out_dir / "sheet.png"))
    cli.emit(args, data, "\n".join([*data["files"], *([data["sheet"]] if args.sheet else [])]))
    return cli.OK


def _handle_doctor(args) -> int:
    report = doctor(Path(args.deck_dir) if args.deck_dir else None)
    ok = bool(report["soffice"] and report["pdftoppm"] and all(report["fonts"].values()))
    cli.emit(args, {**report, "ok": ok}, "\n".join(_doctor_lines(report)))
    return cli.OK if ok else cli.FINDINGS


def register(subparsers) -> None:
    rnd = subparsers.add_parser("render", help="Render slides to PNG with LibreOffice and poppler.")
    rnd.add_argument("deck", help="deck (.pptx)")
    rnd.add_argument("--slides", help="comma-separated slide numbers, default all")
    rnd.add_argument("--dpi", type=int, default=120, help="resolution, default 120")
    rnd.add_argument("--sheet", action="store_true", help="also write sheet.png, a contact sheet of the rendered slides")
    rnd.add_argument("--out", help="output folder, default <deck-dir>/render or <deck>-render")
    rnd.add_argument("--json", action="store_true", help="machine-readable output")
    rnd.set_defaults(handler=_handle_render)

    doc = subparsers.add_parser("doctor", help="Check that rendering tools and the deck's fonts are installed.")
    doc.add_argument("deck_dir", nargs="?", help="deck folder whose theme fonts to check")
    doc.add_argument("--json", action="store_true", help="machine-readable output")
    doc.set_defaults(handler=_handle_doctor)
