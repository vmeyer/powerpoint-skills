"""Slide-level operations on an open presentation."""

from lxml import etree
from pptx.oxml.ns import qn

from .errors import PptxSkillsError

_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def layout_named(prs, name: str):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    available = ", ".join(f'"{layout.name}"' for layout in prs.slide_layouts)
    raise PptxSkillsError(f'unknown layout "{name}"; available layouts: {available}')


def resolve_slide(prs, ref, last=None):
    """A slide from `int` (1-based), "id:<n>" or "$last" (the slide created by the latest add/clone op)."""
    slides = list(prs.slides)
    if ref == "$last":
        if last is None:
            raise PptxSkillsError('"$last" refers to a slide created earlier in this change set, but none was')
        return last
    if isinstance(ref, str) and ref.startswith("id:"):
        text = ref[3:]
        for slide in slides:
            if text.isdigit() and slide.slide_id == int(text):
                return slide
        raise PptxSkillsError(f"no slide with id {text}; slide ids: {', '.join(str(s.slide_id) for s in slides)}")
    if isinstance(ref, int) and not isinstance(ref, bool):
        if 1 <= ref <= len(slides):
            return slides[ref - 1]
        raise PptxSkillsError(f"slide {ref} does not exist; the deck has {len(slides)} slide(s)")
    raise PptxSkillsError(f'invalid slide reference {ref!r}; use a number, "id:<n>" or "$last"')


def _section_lists(prs) -> list:
    """The `p14:sldIdLst` element of every section, in section order."""
    root = prs.part._element
    lists = []
    for section in root.iter(f"{{{_P14}}}section"):
        id_list = section.find(f"{{{_P14}}}sldIdLst")
        if id_list is None:
            id_list = etree.SubElement(section, f"{{{_P14}}}sldIdLst")
        lists.append(id_list)
    return lists


def place_in_section(prs, slide_id: int, previous_id: int | None) -> None:
    """Keep every slide in exactly one section: put `slide_id` into the section of the slide it follows.

    `previous_id` is the id of the slide now directly before it, or None when it is first. A first slide goes
    to the start of the first section; a slide alone in the deck goes to the last section. Decks without
    sections are left alone. Reused by every operation that inserts or moves slides.
    """
    lists = _section_lists(prs)
    if not lists:
        return
    for id_list in lists:
        for ref in [r for r in id_list if r.get("id") == str(slide_id)]:
            id_list.remove(ref)
    new = etree.Element(f"{{{_P14}}}sldId")
    new.set("id", str(slide_id))
    others = [s for s in prs.slides._sldIdLst if s.id != slide_id]
    if previous_id is None:
        if others:
            lists[0].insert(0, new)
        else:
            lists[-1].append(new)
        return
    for id_list in lists:
        for ref in id_list:
            if ref.get("id") == str(previous_id):
                ref.addnext(new)
                return
    lists[-1].append(new)  # the predecessor is in no section: keep the slide listed anyway


def _place(prs, sld_id, index: int) -> None:
    """Move the `p:sldId` element to `index` (0-based) and update its section."""
    id_list = prs.slides._sldIdLst
    id_list.remove(sld_id)
    id_list.insert(index, sld_id)
    previous = id_list[index - 1].id if index > 0 else None
    place_in_section(prs, sld_id.id, previous)


def add_slide(prs, layout_name: str, after=None):
    """Add a slide with the named layout after `after`; at the end when `after` is None."""
    layout = layout_named(prs, layout_name)
    id_list = prs.slides._sldIdLst
    index = len(id_list)
    if after is not None:
        index = [s.id for s in id_list].index(after.slide_id) + 1
    slide = prs.slides.add_slide(layout)
    _place(prs, id_list[-1], index)
    return slide


def move_slide(prs, slide, to: int) -> None:
    """Move the slide so that it ends up at 1-based position `to`."""
    id_list = prs.slides._sldIdLst
    count = len(id_list)
    if isinstance(to, bool) or not isinstance(to, int) or not 1 <= to <= count:
        raise PptxSkillsError(f"cannot move to position {to!r}; the deck has {count} slide(s)")
    sld_id = next(s for s in id_list if s.id == slide.slide_id)
    if list(id_list).index(sld_id) == to - 1:
        return
    _place(prs, sld_id, to - 1)


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
