"""Slide-level operations on an open presentation."""


def delete_slide(prs, slide) -> None:
    """Remove the slide from the slide list and drop its relationship so the part is not saved."""
    id_list = prs.slides._sldIdLst
    for sld_id in list(id_list):
        if sld_id.id == slide.slide_id:
            prs.part.drop_rel(sld_id.rId)
            id_list.remove(sld_id)
            return
