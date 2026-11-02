"""Animation timing (`p:timing`) kept free of references to shapes or paragraphs that are gone (R27).

PowerPoint asks to repair a slide whose animations target a missing shape id or paragraph.
"""

from pptx.oxml.ns import qn

# time nodes: containers (par, seq, excl) and behaviours; removing one removes one animation step
_TIME_NODES = {qn(f"p:{t}") for t in (
    "par", "seq", "excl", "set", "anim", "animClr", "animEffect", "animMotion", "animRot", "animScale", "cmd", "audio", "video",
)}
_CONTAINERS = {qn(f"p:{t}") for t in ("par", "seq", "excl")}
_BUILDS = {qn(f"p:{t}") for t in ("bldP", "bldDgm", "bldOleChart", "bldGraphic")}


def _node_of(target):
    """The time node an animation target belongs to: the behaviour it drives or the sequence it triggers."""
    return next((a for a in target.iterancestors() if a.tag in _TIME_NODES), None)


def _remove(timing, node) -> None:
    """Remove `node` and every container left without children above it, and the builds of shapes that are no
    longer animated at all; drop `p:timing` when nothing is left."""
    parent = node.getparent()
    if parent is None or timing not in list(parent.iterancestors()) + [parent]:
        return  # already gone with an earlier node
    affected = {t.get("spid") for t in node.iter(qn("p:spTgt"))}
    parent.remove(node)
    while parent.tag == qn("p:childTnLst") and len(parent) == 0:
        owner = parent.getparent().getparent()  # childTnLst -> cTn -> par/seq/excl
        if owner is None or owner.tag not in _CONTAINERS:
            break
        parent = owner.getparent()
        parent.remove(owner)
    still = {t.get("spid") for t in timing.iter(qn("p:spTgt"))}
    for build in [b for b in timing.iter() if b.tag in _BUILDS and b.get("spid") in affected - still]:
        build.getparent().remove(build)
    _prune(timing)


def _prune(timing) -> None:
    for tag in ("p:tnLst", "p:bldLst"):
        el = timing.find(qn(tag))
        if el is not None and len(el) == 0:
            timing.remove(el)
    if timing.find(qn("p:tnLst")) is None and timing.find(qn("p:bldLst")) is None and timing.getparent() is not None:
        timing.getparent().remove(timing)


def drop_shapes(sld, shape_ids: set[int]) -> int:
    """Remove every animation step and build that targets one of `shape_ids`; returns how many were removed."""
    timing = sld.find(qn("p:timing"))
    if timing is None:
        return 0
    wanted = {str(i) for i in shape_ids}
    nodes = [_node_of(t) for t in timing.iter(qn("p:spTgt")) if t.get("spid") in wanted]
    builds = [b for b in timing.iter() if b.tag in _BUILDS and b.get("spid") in wanted]
    for el in [n for n in nodes if n is not None] + builds:
        _remove(timing, el)
    return len(nodes) + len(builds)


def drop_paragraphs(sld, shape_id: int, count: int) -> int:
    """Shape `shape_id` now has `count` paragraphs: remove the steps that animate later ones and shorten ranges
    that reach past the end. Returns how many steps were removed or shortened."""
    timing = sld.find(qn("p:timing"))
    if timing is None:
        return 0
    changed = 0
    for target in list(timing.iter(qn("p:spTgt"))):
        rng = target.find(f"{qn('p:txEl')}/{qn('p:pRg')}")
        if target.get("spid") != str(shape_id) or rng is None:
            continue
        if int(rng.get("st", 0)) >= count:
            node = _node_of(target)
            if node is not None:
                _remove(timing, node)
                changed += 1
        elif int(rng.get("end", 0)) >= count:
            rng.set("end", str(count - 1))
            changed += 1
    return changed


def shape_ids(element) -> set[int]:
    """The ids of a shape element and of every shape inside it (group children)."""
    return {int(el.get("id")) for el in element.iter(qn("p:cNvPr")) if (el.get("id") or "").isdigit()}
