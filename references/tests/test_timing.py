"""Animations stay consistent when shapes or paragraphs disappear (R27)."""

import pytest
from lxml import etree
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Cm

from conftest import apply, layout_named
from pptx_skills.checks import run_checks
from pptx_skills.opc import open_presentation


def _effect(node_id: int, spid: int, paragraph: int | None = None) -> str:
    """A click group with one entrance effect on shape `spid` (or one of its paragraphs)."""
    if paragraph is None:
        target = f'<p:spTgt spid="{spid}"/>'
    else:
        target = f'<p:spTgt spid="{spid}"><p:txEl><p:pRg st="{paragraph}" end="{paragraph}"/></p:txEl></p:spTgt>'
    return (
        f'<p:par><p:cTn id="{node_id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
        f'<p:par><p:cTn id="{node_id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        f'<p:par><p:cTn id="{node_id + 2}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" '
        'nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        f'<p:set><p:cBhvr><p:cTn id="{node_id + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        f'<p:tgtEl>{target}</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
        '<p:to><p:strVal val="visible"/></p:to></p:set>'
        "</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>"
    )


def _timing(effects: str, trigger: int, triggered: int, builds: list[str]) -> str:
    """The main click sequence plus an interactive sequence started by clicking shape `trigger`."""
    trigger_cond = f'<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="{trigger}"/></p:tgtEl></p:cond>'
    return (
        f'<p:timing {nsdecls("p")}><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        f"{effects}</p:childTnLst></p:cTn>"
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
        '<p:seq concurrent="1" nextAc="seek"><p:cTn id="90" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" '
        f'nodeType="interactiveSeq"><p:stCondLst>{trigger_cond}</p:stCondLst><p:childTnLst>'
        f'<p:par><p:cTn id="91" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        f'{_effect(92, triggered)}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
        f"<p:nextCondLst>{trigger_cond}</p:nextCondLst></p:seq>"
        "</p:childTnLst></p:cTn></p:par></p:tnLst>"
        f'<p:bldLst>{"".join(builds)}</p:bldLst></p:timing>'
    )


@pytest.fixture
def animated(make_deck):
    """Shapes "badge" (entrance), "list" (three paragraphs built one by one), "cell" in group "row" (the group
    enters), and "button", whose click starts the entrance of "note"."""

    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        shapes = {}
        for name, x in (("badge", 1), ("list", 5), ("button", 10), ("note", 14)):
            shapes[name] = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(2), Cm(3), Cm(3))
            shapes[name].name = name
        shapes["list"].text_frame.text = "one\ntwo\nthree"
        cell = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1), Cm(8), Cm(3), Cm(2))
        cell.name = "cell"
        group = slide.shapes.add_group_shape([cell])
        group.name = "row"
        ids = {name: shape.shape_id for name, shape in {**shapes, "cell": cell, "row": group}.items()}
        effects = _effect(10, ids["badge"]) + "".join(_effect(20 + 10 * p, ids["list"], p) for p in range(3)) + _effect(60, ids["row"])
        builds = [f'<p:bldP spid="{ids["badge"]}" grpId="0" animBg="1"/>', f'<p:bldP spid="{ids["list"]}" grpId="0" build="p"/>',
                  f'<p:bldP spid="{ids["note"]}" grpId="0" animBg="1"/>']
        slide._element.find(qn("p:clrMapOvr")).addnext(etree.fromstring(_timing(effects, ids["button"], ids["note"], builds)))

    return make_deck(build, name="animated.pptx")


def timing(deck):
    return open_presentation(deck).slides[0]._element.find(qn("p:timing"))


def targets(deck) -> list[tuple[int, tuple[int, int] | None]]:
    """(shape id, paragraph range or None) of every animation target, in document order."""
    found = []
    root = timing(deck)
    for tgt in ([] if root is None else root.iter(qn("p:spTgt"))):
        rng = tgt.find(f"{qn('p:txEl')}/{qn('p:pRg')}")
        found.append((int(tgt.get("spid")), None if rng is None else (int(rng.get("st")), int(rng.get("end")))))
    return found


def builds(deck) -> list[int]:
    root = timing(deck)
    return [] if root is None else [int(b.get("spid")) for b in root.iter(qn("p:bldP"))]


def ids(deck) -> dict[str, int]:
    slide = open_presentation(deck).slides[0]
    out = {}
    for shape in slide.shapes:
        out[shape.name] = shape.shape_id
        if shape.shape_type is not None and shape.shape_type.name == "GROUP":
            out.update({child.name: child.shape_id for child in shape.shapes})
    return out


def assert_no_dangling(deck):
    live = set(ids(deck).values())
    assert {spid for spid, _ in targets(deck)} | set(builds(deck)) <= live
    report = run_checks(deck, all_slides=True)
    assert [f for f in report["findings"] if f["category"] == "structure"] == []


def test_delete_shape_removes_its_animation(animated):
    before = ids(animated)
    res = apply(animated, [{"op": "delete_shape", "slide": 1, "target": {"name": "badge"}}])
    assert before["badge"] not in {spid for spid, _ in targets(animated)} and before["badge"] not in builds(animated)
    assert (before["list"], (2, 2)) in targets(animated) and before["list"] in builds(animated)
    assert any('"badge"' in m and "animation" in m for m in res["messages"])
    assert_no_dangling(animated)


def test_deleting_a_trigger_removes_its_interactive_sequence(animated):
    before = ids(animated)
    apply(animated, [{"op": "delete_shape", "slide": 1, "target": {"name": "button"}}])
    assert before["button"] not in {spid for spid, _ in targets(animated)}
    assert before["note"] not in {spid for spid, _ in targets(animated)}  # the effect only that click started
    assert timing(animated).find(f".//{qn('p:cTn')}[@nodeType='interactiveSeq']") is None
    assert_no_dangling(animated)


def test_fewer_paragraphs_drop_their_build_steps(animated):
    list_id = ids(animated)["list"]
    res = apply(animated, [{"op": "set_text", "slide": 1, "target": {"name": "list"}, "text": "one\ntwo"}])
    assert [rng for spid, rng in targets(animated) if spid == list_id] == [(0, 0), (1, 1)]
    assert list_id in builds(animated)
    assert any('"list"' in m and "paragraph" in m for m in res["messages"])
    res = apply(animated, [{"op": "update_shape", "slide": 1, "target": {"name": "list"}, "text": "only"}])
    assert [rng for spid, rng in targets(animated) if spid == list_id] == [(0, 0)]
    assert any('"list"' in m for m in res["messages"])
    res = apply(animated, [{"op": "set_text", "slide": 1, "target": {"name": "list"}, "text": "still one"}])
    assert not any("animation" in m for m in res["messages"])  # nothing to remove
    assert_no_dangling(animated)


def test_a_paragraph_range_that_reaches_past_the_text_is_shortened(make_deck):
    def build(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Blank"))
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1), Cm(1), Cm(5), Cm(5))
        box.name = "box"
        box.text_frame.text = "a\nb\nc"
        effect = _effect(10, box.shape_id, 0).replace('end="0"', 'end="2"')
        xml = _timing(effect, box.shape_id, box.shape_id, [f'<p:bldP spid="{box.shape_id}" grpId="0" build="p"/>'])
        slide._element.find(qn("p:clrMapOvr")).addnext(etree.fromstring(xml))

    deck = make_deck(build, name="range.pptx")
    box_id = ids(deck)["box"]
    apply(deck, [{"op": "set_text", "slide": 1, "target": {"name": "box"}, "text": "a\nb"}])
    assert (box_id, (0, 1)) in targets(deck)


def test_ungroup_removes_the_group_animation(animated):
    row = ids(animated)["row"]
    res = apply(animated, [{"op": "ungroup", "slide": 1, "target": {"name": "row"}}])
    assert row not in {spid for spid, _ in targets(animated)}
    assert any('"row"' in m for m in res["messages"])
    assert_no_dangling(animated)


def test_timing_goes_when_nothing_is_animated_any_more(animated):
    apply(animated, [{"op": "delete_shape", "slide": 1, "target": {"name": name}} for name in ("badge", "list", "row", "button")])
    assert timing(animated) is None
    assert_no_dangling(animated)


def test_deleting_a_group_removes_the_animation_of_its_children(animated):
    cell = ids(animated)["cell"]  # animate the child instead of the group
    xml = etree.tostring(timing(animated)).decode().replace(f'spid="{ids(animated)["row"]}"', f'spid="{cell}"')
    prs = open_presentation(animated)
    sld = prs.slides[0]._element
    sld.replace(sld.find(qn("p:timing")), etree.fromstring(xml))
    prs.save(str(animated))
    apply(animated, [{"op": "delete_shape", "slide": 1, "target": {"name": "row"}}])
    assert cell not in {spid for spid, _ in targets(animated)}
    assert_no_dangling(animated)
