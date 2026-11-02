"""Slide-level operations on an open presentation."""

import copy
import hashlib

from lxml import etree
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import PartFactory
from pptx.oxml.ns import qn
from pptx.parts.chart import ChartPart

from .errors import PptxSkillsError

_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
# markup-compatibility attributes whose values are lists of namespace prefixes
_PREFIX_LISTS = ("Requires", f"{{{_MC}}}Ignorable", f"{{{_MC}}}MustUnderstand")
FOOTER_TYPES = (PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER)
_HYPERLINKS = (qn("a:hlinkClick"), qn("a:hlinkHover"), qn("a:hlinkMouseOver"))
# elements whose only job is to point at a slide: hyperlinks, and `p:sld` in outline views and custom shows
_SLIDE_POINTERS = (*_HYPERLINKS, qn("p:sld"))


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


def refresh_slide_numbers(prs) -> None:
    """Write each slide's number into its slide-number fields, the cached text PowerPoint stores on save.

    The field stays a field; the cached text is what `check` measures, so it must be the number, not the
    layout's prompt.
    """
    for number, slide in enumerate(prs.slides, 1):
        for field in slide.shapes._spTree.iter(qn("a:fld")):
            text = field.find(qn("a:t"))
            if field.get("type") == "slidenum" and text is not None:
                text.text = str(number)


def _renumber_slide_parts(prs) -> None:
    """Name the slide parts slide1.xml … slideN.xml in slide order again, after a slide part was removed.

    python-pptx names a new slide part after the slide count, so without this the next added slide would
    take a name still in use, and a package with two parts of one name is broken.
    """
    prs.part.rename_slide_parts([sld_id.rId for sld_id in prs.slides._sldIdLst])


def _place(prs, sld_id, index: int) -> None:
    """Move the `p:sldId` element to `index` (0-based), update its section and the slide numbers."""
    id_list = prs.slides._sldIdLst
    id_list.remove(sld_id)
    id_list.insert(index, sld_id)
    previous = id_list[index - 1].id if index > 0 else None
    place_in_section(prs, sld_id.id, previous)
    refresh_slide_numbers(prs)


def copy_footer_placeholders(slide, layout) -> None:
    """Put the layout's date, footer and slide-number placeholders on the slide, with their text and fields.

    python-pptx leaves them out of a new slide; PowerPoint adds them when header and footer are on, and the
    master's footer, number and date must appear on every slide. Position and formatting stay inherited.
    """
    present = {ph.placeholder_format.idx for ph in slide.placeholders}
    for source in layout.placeholders:
        idx = source.placeholder_format.idx
        if source.placeholder_format.type not in FOOTER_TYPES or idx in present:
            continue
        slide.shapes.clone_placeholder(source)
        body = source._element.find(qn("p:txBody"))
        if body is None:
            continue
        body = copy.deepcopy(body)
        for tag in ("a:bodyPr", "a:lstStyle"):  # inherited from the layout; the slide keeps them empty
            old = body.find(qn(tag))
            if old is not None:
                old.addprevious(etree.Element(qn(tag)))
                body.remove(old)
        slide.placeholders[idx]._element.append(body)


def add_slide(prs, layout_name: str, after=None):
    """Add a slide with the named layout after `after`; at the end when `after` is None."""
    layout = layout_named(prs, layout_name)
    id_list = prs.slides._sldIdLst
    index = len(id_list)
    if after is not None:
        index = [s.id for s in id_list].index(after.slide_id) + 1
    slide = prs.slides.add_slide(layout)
    copy_footer_placeholders(slide, layout)
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


def _xml(part):
    """(root element, store) of a part: python-pptx keeps XML parts parsed, other parts only as bytes."""
    element = getattr(part, "_element", None)
    if element is not None:
        return element, lambda: None
    root = etree.fromstring(part.blob)

    def store():
        part.blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    return root, store


def _links_to(prs, part) -> list:
    """(part, rId) of every relationship to the slide `part` except the presentation's own.

    Only parts reachable without passing through `part` count: its own notes slide links back to it.
    """
    found, seen, todo = [], {part, prs.part}, [prs.part]
    while todo:
        source = todo.pop()
        for rid, rel in source.rels.items():
            if rel.is_external:
                continue
            if rel.target_part is part:
                if source is not prs.part:
                    found.append((source, rid))
            elif rel.target_part not in seen:
                seen.add(rel.target_part)
                todo.append(rel.target_part)
    return found


def _choice_of(el):
    return next(el.iterancestors(f"{{{_MC}}}Choice"), None)


def _unlink(prs, slide, notes: list[str]) -> None:
    """Remove every link to `slide` from the rest of the deck and drop the relationships behind them.

    Hyperlinks and outline entries are removed; a zoom (markup inside `mc:Choice`) becomes its fallback,
    usually a picture, as in applications without zoom support. Raises before any change when a link is
    of another kind.
    """
    links = [(source, rid, *_xml(source)) for source, rid in _links_to(prs, slide.part)]
    for source, rid, root, _ in links:
        for el, _, value in _refs(root):
            if value == rid and el.tag not in _SLIDE_POINTERS and _choice_of(el) is None:
                name = etree.QName(el).localname
                raise PptxSkillsError(f"cannot delete slide id {slide.slide_id}: {source.partname} refers to it in <{name}>, which cannot be removed automatically")
    for source, rid, root, store in links:
        while refs := [el for el, _, value in _refs(root) if value == rid]:
            choice = _choice_of(refs[0])
            if choice is None:
                refs[0].getparent().remove(refs[0])
                continue
            alternate = choice.getparent()
            fallback = alternate.find(f"{{{_MC}}}Fallback")
            for child in list(fallback) if fallback is not None else []:
                alternate.addprevious(child)
            alternate.getparent().remove(alternate)
        store()
        source.rels.pop(rid)
    ids = {s.part: s.slide_id for s in prs.slides}
    linking = sorted({ids[source] for source, *_ in links if source in ids})
    if linking:
        notes.append(f"deleted slide id {slide.slide_id}: the links to it on slide id(s) {', '.join(map(str, linking))} were removed")


def _repoint_links(prs, old, new) -> None:
    """Point every link to the slide part `old` (hyperlinks, zooms, outline entries) at `new`."""
    for source, rid in _links_to(prs, old):
        root, store = _xml(source)
        new_rid = source.relate_to(new, source.rels[rid].reltype)
        for el, attr, value in _refs(root):
            if value == rid:
                el.set(attr, new_rid)
        store()
        source.rels.pop(rid)


def delete_slide(prs, slide, notes: list[str] | None = None) -> None:
    """Remove the slide from the slide list, sections and custom shows, and drop its relationship.

    Links to it from other slides (hyperlinks, zooms) and from the outline view are removed first, so the
    old slide part does not stay in the file; a note for the user is appended to `notes`.
    """
    root = prs.part._element
    id_list = prs.slides._sldIdLst
    for sld_id in list(id_list):
        if sld_id.id != slide.slide_id:
            continue
        _unlink(prs, slide, notes if notes is not None else [])
        for ref in root.iter(f"{{{_P14}}}sldId"):  # sections keep their (possibly empty) entries
            if ref.get("id") == str(sld_id.id):
                ref.getparent().remove(ref)
        for ref in list(root.iter(qn("p:sld"))):  # custom shows reference the relationship id
            if ref.get(qn("r:id")) == sld_id.rId:
                ref.getparent().remove(ref)
        prs.part.drop_rel(sld_id.rId)
        id_list.remove(sld_id)
        _renumber_slide_parts(prs)
        refresh_slide_numbers(prs)
        return


# --- copying slides between decks -------------------------------------------------------------

def _layout_for(prs, source):
    """The layout of `prs` named like the `source` layout: the same layout, else one under an equally named master."""
    found = [layout for master in prs.slide_masters for layout in master.slide_layouts if layout.name == source.name]
    if not found:
        names = dict.fromkeys(layout.name for master in prs.slide_masters for layout in master.slide_layouts)
        available = ", ".join(f'"{name}"' for name in names)
        raise PptxSkillsError(f"layout '{source.name}' not in working copy; available: {available}")
    for match in (lambda layout: layout.part is source.part, lambda layout: layout.slide_master.name == source.slide_master.name):
        hit = next((layout for layout in found if match(layout)), None)
        if hit is not None:
            return hit
    return found[0]


def _refs(root) -> list:
    """(element, attribute, rId) for every relationship reference in the subtree."""
    return [
        (el, name, value)
        for el in root.iter()
        if isinstance(el.tag, str)
        for name, value in el.attrib.items()
        if name.startswith(f"{{{_R}}}") and value
    ]


def _type_name(rel) -> str:
    name = rel.reltype.rsplit("/", 1)[-1]
    return "oleObject" if name == "package" else name  # an embedded Office file is an OLE object on a slide


def _copyable(rel) -> bool:
    if rel.reltype == RT.IMAGE:
        return True
    if rel.reltype == RT.HYPERLINK:
        return rel.is_external
    return rel.reltype in (RT.SLIDE, RT.CHART) and not rel.is_external


def _sha1(part) -> str:
    return hashlib.sha1(part.blob).hexdigest()


class _SlideCopy:
    """Copies one slide into a destination deck and re-creates every relationship its XML refers to.

    Pictures (also SVG) are shared by content, charts get their own part and workbook, external links are
    kept, links to other slides only within the same deck, PowerPoint tags are dropped. Anything else
    (embedded objects, media, diagrams) is refused before the destination changes.
    """

    def __init__(self, src_prs, src_slide, dst_prs, notes: list[str] | None):
        self.src, self.dst = src_slide, dst_prs
        self.number = [s.slide_id for s in src_prs.slides].index(src_slide.slide_id) + 1
        self.notes = notes if notes is not None else []
        self.package = dst_prs.part.package
        self.same_deck = src_prs.part.package is self.package
        self._images = None  # sha1 -> image part of the destination, filled on first use

    def make(self):
        """Append the copy to the slide list (no section yet) and return it."""
        layout = _layout_for(self.dst, self.src.slide_layout)
        src = self.src._element
        c_sld = src.find(qn("p:cSld"))
        roots = [copy.deepcopy(el) for el in (c_sld.find(qn("p:spTree")), c_sld.find(qn("p:bg"))) if el is not None]
        self._drop_tags(roots, c_sld)
        for root in roots:
            self._check(self.src.part, root)

        slide = self.dst.slides.add_slide(layout)
        sld = slide._element
        tree = sld.find(qn("p:cSld")).find(qn("p:spTree"))
        for child in list(tree):  # the placeholders python-pptx copied from the layout
            tree.remove(child)
        tree.extend(list(roots[0]))
        if len(roots) > 1:
            tree.addprevious(roots[1])
        color_map = src.find(qn("p:clrMapOvr"))
        if color_map is not None:
            old = sld.find(qn("p:clrMapOvr"))
            old.addprevious(copy.deepcopy(color_map))
            sld.remove(old)
        if src.get("showMasterSp") is not None:
            sld.set("showMasterSp", src.get("showMasterSp"))
        self._relink(tree, self.src.part, slide.part)
        if len(roots) > 1:
            self._relink(roots[1], self.src.part, slide.part)
        self._declare_prefixes(sld, c_sld)
        self._copy_notes(slide)
        return slide

    def _refuse(self, what: str):
        raise PptxSkillsError(f"slide {self.number} uses {what}, which cannot be cloned")

    def _drop_tags(self, roots, c_sld) -> None:
        lists = [el for root in roots for el in root.iter(qn("p:custDataLst"))]
        for el in lists:
            el.getparent().remove(el)
        if lists or c_sld.find(qn("p:custDataLst")) is not None:
            self._note(f"slide {self.number}: its PowerPoint tags (add-in data) were not copied")

    def _note(self, text: str) -> None:
        if text not in self.notes:
            self.notes.append(text)

    def _check(self, part, root) -> None:
        for el, _, rid in _refs(root):
            rel = part.rels.get(rid)
            if rel is None:
                raise PptxSkillsError(f"slide {self.number} refers to relationship {rid}, which its deck does not have")
            if not _copyable(rel):
                self._refuse(_type_name(rel))
            if rel.reltype == RT.SLIDE and not self.same_deck and el.tag not in _HYPERLINKS:
                # only a hyperlink can be dropped; a zoom without its slide would be an empty frame
                self._refuse("a Slide Zoom" if _choice_of(el) is not None else f"a link in <{etree.QName(el).localname}> to another slide")
            if rel.reltype == RT.CHART:
                self._check_chart(rel.target_part)

    def _check_chart(self, chart) -> None:
        data = chart._element.find(qn("c:externalData"))
        for el, _, rid in _refs(chart._element):
            rel = chart.rels.get(rid)
            if el is data:
                embedded = rel is not None and not rel.is_external and rel.reltype == RT.PACKAGE
                if not embedded or rel.target_part.content_type != CT.SML_SHEET:
                    self._refuse("a chart whose data is not an embedded .xlsx workbook")
            elif rel is None or rel.reltype != RT.IMAGE:
                self._refuse(f"a chart with {_type_name(rel) if rel is not None else 'a missing part'}")

    def _relink(self, root, src_part, dst_part) -> None:
        """Point every reference in `root` (already in `dst_part`) at a relationship of `dst_part`."""
        new_ids = {}
        for el, attr, rid in _refs(root):
            if rid not in new_ids:
                new_ids[rid] = self._relate(src_part.rels[rid], dst_part)
            if new_ids[rid] is None:
                el.getparent().remove(el)  # a link to another slide that does not exist here
            else:
                el.set(attr, new_ids[rid])

    def _relate(self, rel, dst_part) -> str | None:
        if rel.reltype == RT.IMAGE:
            if rel.is_external:
                return dst_part.relate_to(rel.target_ref, RT.IMAGE, is_external=True)
            return dst_part.relate_to(self._image(rel.target_part), RT.IMAGE)
        if rel.reltype == RT.HYPERLINK:
            return dst_part.relate_to(rel.target_ref, RT.HYPERLINK, is_external=True)
        if rel.reltype == RT.SLIDE:
            if self.same_deck:
                return dst_part.relate_to(rel.target_part, RT.SLIDE)
            self._note(f"slide {self.number}: a link to another slide was removed; it only works inside its own deck")
            return None
        return self._chart(rel.target_part, dst_part)

    def _image(self, part):
        """The destination's image part with the same bytes, created when there is none."""
        if self._images is None:
            self._images = {}
            for rel in self.package.iter_rels():
                if rel.reltype == RT.IMAGE and not rel.is_external:
                    self._images.setdefault(_sha1(rel.target_part), rel.target_part)
        key = _sha1(part)
        if key not in self._images:
            name = self.package.next_image_partname(part.partname.ext)
            self._images[key] = PartFactory(name, part.content_type, self.package, part.blob)
        return self._images[key]

    def _chart(self, source, dst_part) -> str:
        """A new chart part with the source's XML and a copy of its embedded workbook."""
        root = etree.fromstring(source.blob)
        data = root.find(qn("c:externalData"))
        workbook = None
        if data is not None:
            workbook = source.related_part(data.get(qn("r:id"))).blob
            root.remove(data)
        name = self.package.next_partname(ChartPart.partname_template)
        chart = ChartPart.load(name, CT.DML_CHART, self.package, etree.tostring(root))
        rid = dst_part.relate_to(chart, RT.CHART)  # reachable first, so the parts below get free names
        self._relink(chart._element, source, chart)
        if workbook is not None:
            chart.chart_workbook.update_from_xlsx_blob(workbook)
        return rid

    def _declare_prefixes(self, sld, c_sld) -> None:
        """Declare on the slide root the prefixes that copied markup-compatibility attributes name in their values.

        lxml carries the namespaces of element and attribute names along with a copy, but not those only
        named inside a value such as `mc:Choice Requires="a14"`; without them PowerPoint asks to repair the file.
        """
        wanted = {}
        for el in c_sld.iter():
            if not isinstance(el.tag, str):
                continue
            for attr in _PREFIX_LISTS:
                for prefix in (el.get(attr) or "").split():
                    if prefix in el.nsmap and prefix not in sld.nsmap:
                        wanted[prefix] = el.nsmap[prefix]
        if wanted:
            keep = [*wanted, *(prefix for prefix in sld.nsmap if prefix)]
            etree.cleanup_namespaces(sld, top_nsmap=wanted, keep_ns_prefixes=keep)

    def _copy_notes(self, slide) -> None:
        if not self.src.has_notes_slide:
            return
        frame = self.src.notes_slide.notes_text_frame
        if frame is None or not frame.text.strip():
            return
        target = slide.notes_slide.notes_text_frame
        if target is not None:
            target.text = frame.text


def clone_slide(src_prs, src_slide, dst_prs, after=None, notes: list[str] | None = None):
    """Copy `src_slide` of `src_prs` into `dst_prs` after `after` (at the end when None); return the copy.

    The copy uses the destination layout of the same name and keeps shapes, background, pictures, charts
    with their data, external links and the notes text. It joins the section of the slide before it.
    Messages for the user (dropped tags or links) are appended to `notes`. Raises before `dst_prs` changes
    when the layout is missing or the slide uses parts that cannot be copied. `src_prs` may be `dst_prs`.
    """
    slide = _SlideCopy(src_prs, src_slide, dst_prs, notes).make()
    id_list = dst_prs.slides._sldIdLst
    index = len(id_list) - 1
    if after is not None:
        index = [s.id for s in id_list].index(after.slide_id) + 1
    _place(dst_prs, id_list[-1], index)
    return slide


def replace_slide(dst_prs, target, src_prs, src_slide, notes: list[str] | None = None):
    """Put a copy of `src_slide` in place of `target`: same position, slide id, sections and custom shows.

    Links to the target from other slides, zooms and the outline view now lead to the copy.
    """
    slide = _SlideCopy(src_prs, src_slide, dst_prs, notes).make()
    id_list = dst_prs.slides._sldIdLst
    new = id_list[-1]
    old = next(s for s in id_list if s.id == target.slide_id)
    old_rid = old.rId
    for ref in dst_prs.part._element.iter(qn("p:sld")):  # custom shows reference the relationship id
        if ref.get(qn("r:id")) == old_rid:
            ref.set(qn("r:id"), new.rId)
    old.rId = new.rId
    id_list.remove(new)
    _repoint_links(dst_prs, target.part, slide.part)
    dst_prs.part.drop_rel(old_rid)
    _renumber_slide_parts(dst_prs)
    refresh_slide_numbers(dst_prs)
    return slide
