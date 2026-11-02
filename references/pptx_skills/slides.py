"""Slide-level operations on an open presentation."""

from pptx.oxml.ns import qn

_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def delete_slide(prs, slide) -> None:
    """Remove the slide from the slide list, sections and custom shows, and drop its relationship."""
    root = prs.part._element
    id_list = prs.slides._sldIdLst
    for sld_id in list(id_list):
        if sld_id.id != slide.slide_id:
            continue
        for ref in root.iter(f"{{{_P14}}}sldId"):  # sections keep their (possibly empty) entries
            if ref.get("id") == str(sld_id.id):
                ref.getparent().remove(ref)
        for ref in list(root.iter(qn("p:sld"))):  # custom shows reference the relationship id
            if ref.get(qn("r:id")) == sld_id.rId:
                ref.getparent().remove(ref)
        prs.part.drop_rel(sld_id.rId)
        id_list.remove(sld_id)
        return
