import io
import zipfile

import pytest
from pptx import Presentation

_MAIN = b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
_TEMPLATE = b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"


@pytest.fixture
def default_pptx(tmp_path):
    path = tmp_path / "default.pptx"
    Presentation().save(str(path))
    return path


@pytest.fixture
def potx_file(tmp_path, default_pptx):
    path = tmp_path / "template.potx"
    with zipfile.ZipFile(default_pptx) as src, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(_MAIN, _TEMPLATE)
            dst.writestr(item, data)
    return path


@pytest.fixture
def make_deck(tmp_path):
    """Factory: make_deck(builder, name="deck.pptx") builds on the default template and returns the path."""

    def factory(builder=None, name="deck.pptx"):
        prs = Presentation()
        if builder:
            builder(prs)
        path = tmp_path / name
        prs.save(str(path))
        return path

    return factory
