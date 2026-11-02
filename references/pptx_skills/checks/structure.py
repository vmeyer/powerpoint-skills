"""File-structure rules: the package stays valid and PowerPoint opens it without repair."""

import io
from collections import Counter

from pptx import Presentation
from pptx.oxml.ns import qn

from . import rule

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_R_ATTRS = tuple(f"{{{_R_NS}}}{name}" for name in ("id", "embed", "link"))


@rule("reopen", "structure", "error", scope="deck")
def reopen(ctx, _):
    expected = len(ctx.prs.slides)
    try:
        buffer = io.BytesIO()
        ctx.prs.save(buffer)
        buffer.seek(0)
        actual = len(Presentation(buffer).slides)
    except Exception as exc:  # any failure means the saved file is unusable
        return [ctx.finding(None, None, f"the saved file cannot be reopened: {exc}", "undo the last change set or restore a snapshot")]
    if actual != expected:
        return [ctx.finding(None, None, f"the saved file has {actual} slides, expected {expected}", "restore a snapshot and apply the change again")]
    return []


@rule("relationships", "structure", "error")
def relationships(ctx, slide):
    rels = slide.pptx.part.rels
    missing = sorted(
        {el.get(attr) for el in slide.pptx._element.iter() for attr in _R_ATTRS if el.get(attr) and el.get(attr) not in rels}
    )
    if not missing:
        return []
    return [ctx.finding(slide, None, f"the slide refers to missing relationships: {', '.join(missing)}", "remove the shape that uses them or add the missing picture, link or media again")]


@rule("duplicate-ids", "structure", "error")
def duplicate_ids(ctx, slide):
    out = []
    ids = Counter(el.get("id") for el in slide.pptx._element.iter(qn("p:cNvPr")))
    for shape_id, count in ids.items():
        if count > 1:
            out.append(ctx.finding(slide, None, f"shape id {shape_id} is used {count} times on this slide", "remove the duplicate shapes and add them again with change sets"))
    sld_ids = Counter(el.get("id") for el in ctx.prs.part._element.iter(qn("p:sldId")))
    if sld_ids[str(slide.id)] > 1:
        out.append(ctx.finding(slide, None, f"slide id {slide.id} is used by several slides", 'delete the duplicate slide and duplicate it again with a clone_slide op from "self"'))
    return out
