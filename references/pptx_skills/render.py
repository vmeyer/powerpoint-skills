"""Render slides to PNG (PowerPoint or LibreOffice, then poppler), contact sheets of layouts and examples, and `doctor`."""

import hashlib
import os
import plistlib
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
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
_POWERPOINT_PATHS = (
    Path("/Applications/Microsoft PowerPoint.app"),
    Path.home() / "Applications/Microsoft PowerPoint.app",
)
_MACOS = sys.platform == "darwin"
RENDERERS = ("auto", "powerpoint", "libreoffice")
RENDERER_ENV = "PPTX_SKILLS_RENDERER"
_LABELS = {"powerpoint": "PowerPoint", "libreoffice": "LibreOffice"}
_TIMEOUT_S = 180
_POWERPOINT_TIMEOUT_S = 120
_CLOSE_TIMEOUT_S = 15
# LibreOffice leaves hidden slides out of a PDF by default, so page n would no longer be slide n
_PDF_FILTER = 'pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}'
_GAP = 16
_LABEL_H = 34
_INSTALL_HINT = (
    "install LibreOffice and poppler: `brew install --cask libreoffice` (add `--appdir=~/Applications` without "
    "admin rights) and `brew install poppler`, or `apt install libreoffice poppler-utils`"
)
_POPPLER_HINT = "install poppler: `brew install poppler` or `apt install poppler-utils`"
# PowerPoint writes only into folders granted one by one in its "Grant Access" dialog (a parent's grant does not
# count), so all renders of a project share one fixed, visible exchange folder, `decks/_export`.
EXPORT_FOLDER = "_export"
_AUTOMATION_HINT = (
    "allow the app that runs `pptx` (your terminal or agent) to control Microsoft PowerPoint under "
    "System Settings → Privacy & Security → Automation"
)
_CANCEL = '"Cancel" makes `pptx` fall back to LibreOffice in auto mode, where LibreOffice is installed'
_POWERPOINT_SETUP = (
    f"PowerPoint needs a one-time setup: {_AUTOMATION_HINT}; and the first PowerPoint export in a project shows "
    f'one "Grant Access" dialog for `decks/{EXPORT_FOLDER}` (for a deck outside a deck folder: `{EXPORT_FOLDER}` '
    f'next to it), so click "Grant Access" once. {_CANCEL}'
)


def _grant_hint(exchange: Path) -> str:
    return (f'PowerPoint is probably showing a "Grant Access" dialog for {exchange}: click "Grant Access" '
            f"(it is asked once per project), then run the command again; {_CANCEL}")
_DENIED = re.compile(r"-1743|not authori[sz]ed|not allowed to send", re.I)
# The copy is opened, exported and closed by its unique name; the user's own presentations are never addressed,
# PowerPoint is never quit, and the paths arrive as arguments rather than as script text.
_EXPORT_SCRIPT = """on run argv
	set deckFile to POSIX file (item 1 of argv)
	set pdfFile to POSIX file (item 2 of argv)
	set stem to item 3 of argv
	tell application "Microsoft PowerPoint"
		with timeout of 110 seconds
			open deckFile
			set deckName to missing value
			repeat 50 times
				repeat with candidate in (get name of every presentation)
					if (candidate as text) starts with stem then
						set deckName to candidate as text
						exit repeat
					end if
				end repeat
				if deckName is not missing value then exit repeat
				delay 0.2
			end repeat
			if deckName is missing value then error "PowerPoint did not open the copy of the deck"
			try
				save presentation deckName in pdfFile as save as PDF
			on error errMsg number errNum
				close presentation deckName saving no
				error errMsg number errNum
			end try
			close presentation deckName saving no
		end timeout
	end tell
end run"""
_CLOSE_SCRIPT = """on run argv
	set stem to item 1 of argv
	if application "Microsoft PowerPoint" is running then
		tell application "Microsoft PowerPoint"
			repeat with candidate in (get name of every presentation)
				if (candidate as text) starts with stem then close presentation (candidate as text) saving no
			end repeat
		end tell
	end if
end run"""
LAYOUTS_SHEET = "layouts.png"
EXAMPLES_SHEET = "examples.png"


class _PowerPointFailed(Exception):
    """PowerPoint produced no PDF; the message says why, `hint` what the user can do about it."""

    def __init__(self, reason: str, hint: str = _POWERPOINT_SETUP):
        super().__init__(reason)
        self.hint = hint


def find_soffice() -> Path | None:
    found = shutil.which("soffice")
    if found:
        return Path(found)
    return next((p for p in _SOFFICE_PATHS if p.is_file()), None)


def find_powerpoint() -> Path | None:
    """The Microsoft PowerPoint app on macOS, None elsewhere or when it is not installed."""
    if not _MACOS:
        return None
    return next((p for p in _POWERPOINT_PATHS if p.is_dir()), None)


def powerpoint_version(app: Path) -> str | None:
    """The version from the app's Info.plist, read without launching PowerPoint."""
    try:
        with open(Path(app) / "Contents/Info.plist", "rb") as handle:
            return plistlib.load(handle).get("CFBundleShortVersionString")
    except (OSError, ValueError, plistlib.InvalidFileException):
        return None


def find_pdftoppm() -> Path | None:
    found = shutil.which("pdftoppm")
    return Path(found) if found else None


def requested_renderer(requested: str | None = None) -> str:
    """The renderer asked for: the `--renderer` option, else PPTX_SKILLS_RENDERER, else auto."""
    value = (requested or os.environ.get(RENDERER_ENV) or "auto").strip().lower()
    if value not in RENDERERS:
        raise PptxSkillsError(f"{RENDERER_ENV} must be one of {', '.join(RENDERERS)}; got {value!r}")
    return value


def select_renderer(requested: str) -> str | None:
    """The installed renderer for a request: auto takes PowerPoint, else LibreOffice; None when none fits."""
    if requested in ("auto", "powerpoint") and find_powerpoint() is not None:
        return "powerpoint"
    if requested in ("auto", "libreoffice") and find_soffice() is not None:
        return "libreoffice"
    return None


def _require_tools(renderer: str | None = None) -> tuple[str, str, Path]:
    """(requested renderer, renderer to use, pdftoppm), or an error that says what to install."""
    want = requested_renderer(renderer)
    chosen, pdftoppm = select_renderer(want), find_pdftoppm()
    problems, hints = [], []
    if chosen is None:
        if want == "powerpoint":
            problems.append(f"PowerPoint not found (asked for by --renderer or {RENDERER_ENV})")
            hints.append("PowerPoint renders on macOS from /Applications or ~/Applications; or use `--renderer libreoffice`")
        else:
            problems.append("LibreOffice (soffice) not found" if want == "libreoffice" else "neither PowerPoint nor LibreOffice found")
            hints.append(_INSTALL_HINT)
    if pdftoppm is None:
        problems.append("poppler (pdftoppm) not found")
        if _INSTALL_HINT not in hints:
            hints.append(_POPPLER_HINT)
    if problems:
        raise PptxSkillsError(f"{' and '.join(problems)}; {'; '.join(hints)}; `pptx doctor` checks the setup")
    return want, chosen, pdftoppm


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


def _libreoffice_pdf(soffice: Path, deck: Path, tmp: Path) -> Path:
    profile = tmp / "lo-profile"
    _run(
        [str(soffice), "--headless", f"-env:UserInstallation={profile.as_uri()}", "--convert-to", _PDF_FILTER,
         "--outdir", str(tmp), str(deck)],
        "LibreOffice conversion",
    )
    pdf = tmp / f"{deck.stem}.pdf"
    if not pdf.is_file():
        raise PptxSkillsError(f"LibreOffice produced no PDF for {deck}")
    return pdf


def _powerpoint_pdf(prs, exchange: Path, pdf: Path) -> Path:
    """Export `prs` to `pdf` through PowerPoint; raises _PowerPointFailed.

    PowerPoint is sandboxed: it writes only into folders the user granted one by one, so `exchange` is the
    project's fixed `_export` folder (see default_export_dir), kept between runs. It receives a copy with a unique
    name and every slide shown (PowerPoint leaves hidden slides out of a PDF, so page n would no longer be slide n);
    the files of this run (copy, PowerPoint's owner file, PDF) are always removed, other files stay.
    """
    stem = f"render-{uuid.uuid4().hex}"
    try:
        exchange.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise _PowerPointFailed(f"cannot create {exchange}: {exc}", "check that the output folder is writable") from exc
    started = done = False
    try:
        for slide in prs.slides:
            slide._element.attrib.pop("show", None)
        copy, exported = exchange / f"{stem}.pptx", exchange / f"{stem}.pdf"
        save_presentation(prs, copy)
        started = True
        try:
            result = subprocess.run(
                ["osascript", "-e", _EXPORT_SCRIPT, str(copy), str(exported), stem],
                capture_output=True, text=True, errors="replace", timeout=_POWERPOINT_TIMEOUT_S,
            )
        except subprocess.TimeoutExpired as exc:
            raise _PowerPointFailed(f"PowerPoint did not answer within {_POWERPOINT_TIMEOUT_S} s",
                                    _grant_hint(exchange)) from exc
        if result.returncode != 0:
            tail = (result.stderr or result.stdout or "").strip()[-400:]
            if _DENIED.search(tail):
                hint = f"To fix it, {_AUTOMATION_HINT}, then run the command again"
            else:
                hint = f"Check the Automation permission ({_AUTOMATION_HINT}). {_grant_hint(exchange)}"
            raise _PowerPointFailed(f"osascript failed (exit {result.returncode}): {tail}", hint)
        if not exported.is_file() or exported.stat().st_size == 0:
            raise _PowerPointFailed(f"PowerPoint wrote no PDF into {exchange}", _grant_hint(exchange))
        shutil.move(str(exported), str(pdf))
        done = True
        return pdf
    except OSError as exc:
        raise _PowerPointFailed(str(exc)) from exc
    finally:
        if started and not done:
            _close_quietly(stem)
        for leftover in exchange.glob(f"*{stem}*"):  # the copy, PowerPoint's ~$ owner file, the PDF
            try:
                leftover.unlink()
            except OSError:
                pass


def _close_quietly(stem: str) -> None:
    """Best effort: close the copy if PowerPoint still has it open, without saving; never launches PowerPoint."""
    try:
        subprocess.run(["osascript", "-e", _CLOSE_SCRIPT, stem], capture_output=True, text=True, errors="replace",
                       timeout=_CLOSE_TIMEOUT_S)
    except (subprocess.TimeoutExpired, OSError):
        pass


def _deck_dir_of_working_copy(deck: Path) -> Path | None:
    from .deckdir import find_deck_dir, load_deck

    deck_dir = find_deck_dir(deck)
    if deck_dir is not None and load_deck(deck_dir)["working_copy"] == Path(deck).name:
        return deck_dir
    return None


def default_out_dir(deck: Path) -> Path:
    deck = Path(deck)
    deck_dir = _deck_dir_of_working_copy(deck)
    return deck_dir / "render" if deck_dir is not None else deck.parent / f"{deck.stem}-render"


def default_export_dir(deck: Path) -> Path:
    """PowerPoint's exchange folder: `decks/_export` for a deck folder's working copy, shared by every deck of the
    project and their contact sheets, else `_export` next to the deck. One folder means one "Grant Access"."""
    deck = Path(deck)
    deck_dir = _deck_dir_of_working_copy(deck)
    return (deck_dir.parent if deck_dir is not None else deck.parent) / EXPORT_FOLDER


def render(
    deck: Path,
    out_dir: Path | None = None,
    slides: list[int] | None = None,
    dpi: int = 120,
    renderer: str | None = None,
    notes: list[str] | None = None,
    export_dir: Path | None = None,
) -> list[Path]:
    """Render the given slides (1-based, default all) to `slide-NN.png` in `out_dir`; returns the files.

    `renderer` is auto, powerpoint or libreoffice (default: PPTX_SKILLS_RENDERER, else auto). When auto picked
    PowerPoint and it fails, LibreOffice renders instead and a message for the user is appended to `notes`.
    PowerPoint exchanges files through `export_dir`, default `default_export_dir(deck)`.
    """
    deck = Path(deck)
    prs = open_presentation(deck)  # also validates the file
    count = len(prs.slides)
    numbers = list(slides) if slides else list(range(1, count + 1))
    bad = [n for n in numbers if not 1 <= n <= count]
    if bad:
        raise PptxSkillsError(f"slide {bad[0]} does not exist; the deck has {count} slide(s)")
    want, chosen, pdftoppm = _require_tools(renderer)
    out_dir = Path(out_dir) if out_dir is not None else default_out_dir(deck)
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise PptxSkillsError(f"cannot create output folder {out_dir}: {exc}") from exc
    with tempfile.TemporaryDirectory(prefix="pptx-render-") as tmp_name:
        tmp = Path(tmp_name)
        pdf = None
        if chosen == "powerpoint":
            try:
                pdf = _powerpoint_pdf(prs, Path(export_dir or default_export_dir(deck)), tmp / "deck.pdf")
            except _PowerPointFailed as exc:
                if want == "powerpoint" or find_soffice() is None:
                    raise PptxSkillsError(f"PowerPoint could not export {deck.name} to PDF: {exc}. {exc.hint}.") from exc
                if notes is not None:
                    notes.append(f"PowerPoint could not render ({exc}), so LibreOffice did; "
                                 "`pptx doctor` shows PowerPoint's one-time setup")
        if pdf is None:
            pdf = _libreoffice_pdf(find_soffice(), deck, tmp)
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


def render_layouts(
    deck: Path, out: Path, renderer: str | None = None, notes: list[str] | None = None, export_dir: Path | None = None
) -> Path:
    """One slide per layout of `deck` (placeholders filled with their role name), labeled with the layout name.

    PowerPoint exchanges files through `export_dir`, default `default_export_dir(deck)`."""
    out = Path(out)
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
        pngs = render(tmp / "layouts.pptx", tmp / "png", dpi=60, renderer=renderer, notes=notes,
                      export_dir=export_dir or default_export_dir(deck))
        return contact_sheet(pngs, names, out)


def render_examples(
    examples: Path, out: Path, renderer: str | None = None, notes: list[str] | None = None, export_dir: Path | None = None
) -> Path:
    out = Path(out)
    with tempfile.TemporaryDirectory(prefix="pptx-examples-") as tmp_name:
        pngs = render(examples, Path(tmp_name), dpi=60, renderer=renderer, notes=notes,
                      export_dir=export_dir or default_export_dir(examples))
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


def update_sheets(
    deck_dir: Path, deck: dict, previous_hash: str | None, force: bool = False, renderer: str | None = None
) -> tuple[str | None, list[str]]:
    """Write style/layouts.png (and examples.png); returns (sources_hash to store, notes for the user).

    Rendering never fails the caller: a missing or failing renderer becomes a note.
    """
    deck_dir = Path(deck_dir)
    style = deck_dir / "style"
    current = sources_hash(deck_dir, deck)
    examples = deck["sources"].get("examples")
    wanted = [style / LAYOUTS_SHEET] + ([style / EXAMPLES_SHEET] if examples else [])
    if not force and previous_hash == current and all(p.is_file() for p in wanted):
        return current, []
    notes: list[str] = []
    # the project's decks/_export, also while init stages the deck in a temporary folder next to the deck folders
    export_dir = deck_dir.resolve().parent / EXPORT_FOLDER
    try:
        _require_tools(renderer)
        render_layouts(deck_dir / deck["working_copy"], style / LAYOUTS_SHEET, renderer, notes, export_dir)
        if examples:
            render_examples(Path(examples), style / EXAMPLES_SHEET, renderer, notes, export_dir)
    except (PptxSkillsError, OSError) as exc:
        return None, [*dict.fromkeys(notes), f"contact sheets skipped: {exc}"]
    return current, list(dict.fromkeys(notes))


def doctor(deck_dir: Path | None = None) -> dict:
    fonts: dict = {}
    brief = load_brief(deck_dir) if deck_dir is not None else None
    for family in dict.fromkeys(((brief or {}).get("theme") or {}).get("fonts", {}).values()):
        if family:
            path = find_font(family)
            fonts[family] = str(path) if path else None
    want = requested_renderer()
    app, soffice, pdftoppm = find_powerpoint(), find_soffice(), find_pdftoppm()
    return {
        "renderer": select_renderer(want),
        "requested_renderer": want,
        "powerpoint": str(app) if app else None,
        "powerpoint_version": powerpoint_version(app) if app else None,
        "soffice": str(soffice) if soffice else None,
        "pdftoppm": str(pdftoppm) if pdftoppm else None,
        "fonts": fonts,
    }


def _doctor_lines(report: dict) -> list[str]:
    want, chosen = report["requested_renderer"], report["renderer"]
    if chosen:
        how = "auto: PowerPoint if installed, else LibreOffice" if want == "auto" else f"set by {RENDERER_ENV}"
        lines = [f"✓ renderer: {_LABELS[chosen]} ({how})"]
    elif want == "auto":
        lines = [f"✗ renderer: none; Microsoft PowerPoint (macOS) or LibreOffice is needed; {_INSTALL_HINT}"]
    else:
        lines = [f"✗ renderer: {_LABELS[want]} is set by {RENDERER_ENV} but not installed"]
    if report["powerpoint"]:
        version = f" {report['powerpoint_version']}" if report["powerpoint_version"] else ""
        lines.append(f"✓ PowerPoint{version}: {report['powerpoint']}")
        lines.append(f"  {_POWERPOINT_SETUP}.")
    else:
        lines.append("· PowerPoint: not found" + ("" if _MACOS else " (renders on macOS only)"))
    if report["soffice"]:
        lines.append(f"✓ soffice: {report['soffice']}" + (" (fallback)" if chosen == "powerpoint" else ""))
    elif chosen == "powerpoint":
        lines.append("· soffice: not found; not needed while PowerPoint renders")
    else:
        lines.append("✗ soffice: not found; brew install --cask libreoffice (add --appdir=~/Applications without admin "
                     "rights) or apt install libreoffice")
    lines.append(f"✓ pdftoppm: {report['pdftoppm']}" if report["pdftoppm"]
                 else "✗ pdftoppm: missing; brew install poppler (or apt install poppler-utils)")
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
    notes: list[str] = []
    files = render(deck, out_dir, slides, args.dpi, args.renderer, notes)
    data = {"files": [str(f) for f in files]}
    if args.sheet:
        numbers = slides or list(range(1, len(files) + 1))
        data["sheet"] = str(render_slide_sheet(files, numbers, out_dir / "sheet.png"))
    if notes:
        data["notes"] = notes
    cli.emit(args, data, "\n".join([*data["files"], *([data["sheet"]] if args.sheet else []), *notes]))
    return cli.OK


def _handle_doctor(args) -> int:
    report = doctor(Path(args.deck_dir) if args.deck_dir else None)
    ok = bool(report["renderer"] and report["pdftoppm"] and all(report["fonts"].values()))
    cli.emit(args, {**report, "ok": ok}, "\n".join(_doctor_lines(report)))
    return cli.OK if ok else cli.FINDINGS


def add_renderer_option(parser) -> None:
    parser.add_argument(
        "--renderer", choices=RENDERERS,
        help=f"auto (default: PowerPoint if installed, else LibreOffice), powerpoint or libreoffice; also {RENDERER_ENV}",
    )


def register(subparsers) -> None:
    rnd = subparsers.add_parser("render", help="Render slides to PNG with PowerPoint or LibreOffice, and poppler.")
    rnd.add_argument("deck", help="deck (.pptx)")
    rnd.add_argument("--slides", help="comma-separated slide numbers, default all")
    rnd.add_argument("--dpi", type=int, default=120, help="resolution, default 120")
    rnd.add_argument("--sheet", action="store_true", help="also write sheet.png, a contact sheet of the rendered slides")
    rnd.add_argument("--out", help="output folder, default <deck-dir>/render or <deck>-render")
    add_renderer_option(rnd)
    rnd.add_argument("--json", action="store_true", help="machine-readable output")
    rnd.set_defaults(handler=_handle_render)

    doc = subparsers.add_parser("doctor", help="Show the renderer and check that rendering tools and the deck's fonts are installed.")
    doc.add_argument("deck_dir", nargs="?", help="deck folder whose theme fonts to check")
    doc.add_argument("--json", action="store_true", help="machine-readable output")
    doc.set_defaults(handler=_handle_doctor)
