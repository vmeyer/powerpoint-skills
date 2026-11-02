"""Opening, saving and locking PowerPoint files."""

import hashlib
import io
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.exc import PackageNotFoundError

from .errors import PptxSkillsError

_MAIN = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
_TEMPLATE = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"


def _potx_as_pptx(path: Path) -> io.BytesIO:
    """python-pptx rejects the template content type; rewrite it in memory."""
    out = io.BytesIO()
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(_TEMPLATE.encode(), _MAIN.encode())
            dst.writestr(item, data)
    out.seek(0)
    return out


def open_presentation(path: Path):
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix not in (".pptx", ".potx"):
        raise PptxSkillsError(f"{path} is not a .pptx or .potx file")
    if not path.is_file():
        raise PptxSkillsError(f"{path} does not exist")
    try:
        if suffix == ".potx":
            return Presentation(_potx_as_pptx(path))
        return Presentation(str(path))
    except (PackageNotFoundError, zipfile.BadZipFile, KeyError) as exc:  # KeyError: valid zip, not a presentation package
        raise PptxSkillsError(f"{path} cannot be read as a presentation: {exc}") from exc


def file_hash(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def save_presentation(prs, path: Path) -> str:
    path = Path(path)
    prs.save(str(path))
    return file_hash(path)


def lock_file(path: Path) -> Path | None:
    """Return PowerPoint's `~$` owner file next to `path` if it exists."""
    path = Path(path)
    lock = path.with_name(f"~${path.name}")
    return lock if lock.exists() else None
