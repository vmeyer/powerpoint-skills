import json

from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Cm, Pt

from conftest import finding, layout_named
from pptx_skills.checks import run_checks
from pptx_skills.cli import main


def add_box(slide, name, text, x=1, y=1, w=20, h=3, size=14):
    box = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    box.name = name
    box.text_frame.word_wrap = True
    box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    paragraphs = text.split("\n")
    box.text_frame.text = paragraphs[0]
    for extra in paragraphs[1:]:
        box.text_frame.add_paragraph().text = extra
    for p in box.text_frame.paragraphs:
        for run in p.runs:
            run.font.size = Pt(size)
    return box


def build(make_deck, fill):
    def builder(prs):
        fill(prs.slides.add_slide(layout_named(prs, "Blank")))

    return make_deck(builder)


def test_overflow_is_error(make_deck):
    deck = build(make_deck, lambda s: add_box(s, "big", " ".join(["word"] * 60), w=5, h=1, size=12))
    assert finding(run_checks(deck), "overflow")["severity"] == "error"


def test_shape_that_grows_does_not_overflow(make_deck):
    def fill(slide):
        box = add_box(slide, "grows", " ".join(["word"] * 60), w=5, h=1, size=12)
        box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    assert finding(run_checks(build(make_deck, fill)), "overflow") is None


def test_fitting_text_is_clean(make_deck):
    deck = build(make_deck, lambda s: add_box(s, "ok", "A short sentence.", w=10, h=2))
    report = run_checks(deck)
    assert report["findings"] == [] and report["summary"] == {"error": 0, "warning": 0, "info": 0}


def test_keyword_chain_only_for_three_items(make_deck):
    texts = {"chain": "Speed · Quality · Scale", "footer": "|   Company AG", "eg": "z. B. Pipelines, Gates, Audit-Trail", "two": "A | B"}

    def fill(slide):
        for i, (name, text) in enumerate(texts.items()):
            add_box(slide, name, text, y=1 + i * 2, h=1.5)

    flagged = {f["shape"] for f in run_checks(build(make_deck, fill))["findings"] if f["rule"] == "keyword-chain"}
    assert flagged == {"chain"}


def test_leftover_prompt_text(kicker_master):
    from pptx import Presentation

    prs = Presentation(str(kicker_master))
    slide = prs.slides.add_slide(layout_named(prs, "Title Only"))
    slide.shapes.title.text = "A full sentence as the takeaway"
    kicker = next(s for s in slide.placeholders if s.placeholder_format.idx == 13)
    kicker.text = "kicker"  # the layout prompt, other case
    prs.save(str(kicker_master))
    assert finding(run_checks(kicker_master), "leftover-text")["shape_id"] == kicker.shape_id


def test_leftover_markers_and_cloned_text(make_deck, tmp_path):
    deck = build(make_deck, lambda s: (add_box(s, "todo", "TODO write this"), add_box(s, "kept", "Example text", y=5)))
    assert finding(run_checks(deck), "leftover-text", "todo")
    assert finding(run_checks(deck), "leftover-text", "kept") is None
    from pptx import Presentation

    prs = Presentation(str(deck))
    slide = prs.slides[0]
    kept = next(s for s in slide.shapes if s.name == "kept")
    folder = tmp_path / "decks/d"
    folder.mkdir(parents=True)
    prs.save(str(folder / "d.pptx"))
    cloned = {str(slide.slide_id): {str(kept.shape_id): "example   TEXT"}}
    (folder / "deck.json").write_text(json.dumps({"working_copy": "d.pptx", "sources": {}, "cloned_texts": cloned}))
    assert finding(run_checks(folder / "d.pptx"), "leftover-text", "kept")


def test_min_size_and_small_zone(make_deck, tmp_path):
    deck = build(make_deck, lambda s: (add_box(s, "body", "Ten points", size=10), add_box(s, "low", "Ten points", y=17, h=1, size=10)))
    assert finding(run_checks(deck), "min-size", "body")["severity"] == "error"
    assert finding(run_checks(deck), "min-size", "low")  # no brief: small zone unknown, 10 pt < 11
    from pptx import Presentation

    folder = tmp_path / "decks/d"
    (folder / "style").mkdir(parents=True)
    Presentation(str(deck)).save(str(folder / "d.pptx"))
    (folder / "deck.json").write_text(json.dumps({"working_copy": "d.pptx", "sources": {}}))
    brief = {"content_area": [1, 4, 22, 12], "rules": {"min_size_pt": 11, "min_size_small_pt": 9}}
    (folder / "style/brief.json").write_text(json.dumps(brief))
    report = run_checks(folder / "d.pptx")
    assert finding(report, "min-size", "body") and finding(report, "min-size", "low") is None


def test_lines_per_paragraph_not_per_shape(make_deck):
    four = build(make_deck, lambda s: add_box(s, "four", "One\nTwo\nThree\nFour", h=6))
    assert finding(run_checks(four), "too-many-lines") is None
    long = build(make_deck, lambda s: add_box(s, "long", " ".join(["word"] * 80), w=6, h=12), ) if False else None
    deck = build(make_deck, lambda s: add_box(s, "long", " ".join(["word"] * 40), w=6, h=12))
    assert finding(run_checks(deck), "too-many-lines")["severity"] == "warning"


def test_empty_placeholder_and_topic_title(make_deck):
    def builder(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Title and Content"))
        slide.shapes.title.text = "Results"

    report = run_checks(make_deck(builder))
    assert finding(report, "empty-placeholder")["severity"] == "warning"
    assert finding(report, "topic-title")["severity"] == "info"


def test_bullets(make_deck):
    def builder(prs):
        slide = prs.slides.add_slide(layout_named(prs, "Title and Content"))
        slide.shapes.title.text = "Six points make a long list here"
        slide.placeholders[1].text = "\n".join(f"Point {i}" for i in range(6))

    assert finding(run_checks(make_deck(builder)), "too-many-bullets")


def test_shrink_to_fit(make_deck):
    from lxml import etree
    from pptx.oxml.ns import qn

    def fill(slide):
        box = add_box(slide, "tight", "Some text", w=10, h=2)
        body = box.text_frame._txBody.find(qn("a:bodyPr"))
        for child in list(body):
            body.remove(child)
        etree.SubElement(body, qn("a:normAutofit")).set("fontScale", "62500")

    report = run_checks(build(make_deck, fill))
    assert finding(report, "shrink-to-fit")["severity"] == "warning"
    assert finding(report, "min-size")  # 14 pt * 0.625 = 8.75 pt


def test_structure_rules(make_deck):
    deck = build(make_deck, lambda s: add_box(s, "ok", "Fine text", h=2))
    from pptx import Presentation

    prs = Presentation(str(deck))
    shapes = prs.slides[0].shapes
    shapes.add_textbox(0, 0, Cm(1), Cm(1))._element.xpath("./p:nvSpPr/p:cNvPr")[0].set("id", str(shapes[0].shape_id))
    prs.save(str(deck))
    report = run_checks(deck)
    assert finding(report, "duplicate-ids")["severity"] == "error"
    assert finding(report, "relationships") is None and finding(report, "reopen") is None


def test_missing_relationship(make_deck):
    deck = build(make_deck, lambda s: add_box(s, "ok", "Fine text", h=2))
    from pptx import Presentation
    from pptx.oxml import parse_xml

    prs = Presentation(str(deck))
    slide = prs.slides[0]
    pic = parse_xml(
        '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<p:nvPicPr><p:cNvPr id="99" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="rId77"/></p:blipFill><p:spPr/></p:pic>'
    )
    slide.shapes._spTree.append(pic)
    prs.save(str(deck))
    assert finding(run_checks(deck), "relationships")["severity"] == "error"


def test_exit_codes_and_report_files(make_deck, tmp_path, capsys):
    bad = build(make_deck, lambda s: add_box(s, "big", " ".join(["word"] * 60), w=5, h=1))
    clean = make_deck(lambda prs: add_box(prs.slides.add_slide(layout_named(prs, "Blank")), "ok", "Fine text", h=2), "clean.pptx")
    assert main(["check", str(bad)]) == 1 and main(["check", str(clean)]) == 0
    out = capsys.readouterr().out
    assert "slide 1 [big] error overflow:" in out and "Checked slides 1:" in out
    assert main(["check", str(bad), "--json"]) == 1
    assert json.loads(capsys.readouterr().out)["summary"]["error"] == 1


def test_deck_folder_default_slides_and_files(deck_dir_from_default):
    from pptx_skills.deckdir import _save_deck, load_deck
    from pptx import Presentation

    work = deck_dir_from_default / "d.pptx"
    prs = Presentation(str(work))
    for text in ("First slide here now", "Second slide here now"):
        prs.slides.add_slide(layout_named(prs, "Title Only")).shapes.title.text = text
    prs.save(str(work))
    deck = load_deck(deck_dir_from_default)
    deck["last_touched"] = [prs.slides[1].slide_id, 9999]
    _save_deck(deck_dir_from_default, deck)
    assert run_checks(deck_dir_from_default)["slides"] == [2]
    assert (deck_dir_from_default / "check/slide-02.json").is_file()
    assert run_checks(deck_dir_from_default, all_slides=True)["slides"] == [1, 2]
    assert run_checks(deck_dir_from_default, slides=[1])["slides"] == [1]


def test_footer_equal_to_layout_footer_is_not_leftover(make_deck):
    import copy

    from pptx.oxml.ns import qn

    def builder(prs):
        layout = layout_named(prs, "Blank")
        layout_footer = next(s for s in layout.placeholders if s.placeholder_format.type == 15)
        layout_footer.text_frame.text = "| Company AG"
        slide = prs.slides.add_slide(layout)
        el = copy.deepcopy(layout_footer._element)
        el.find(qn("p:spPr")).clear()
        slide.shapes._spTree.append(el)

    report = run_checks(make_deck(builder))
    assert finding(report, "leftover-text") is None


def test_short_prefix_is_not_a_prompt():
    from pptx_skills.checks.text import _is_prompt

    prompt = "Click to edit Master title style"
    assert not _is_prompt("Click", prompt) and not _is_prompt("C", prompt)
    assert _is_prompt("Click to edit", prompt)
    assert _is_prompt("Enter Chapter Title Here", "Enter chapter title")


def growing_box(slide, y=2.0, lines=4, name="grows"):
    """A word-wrapped text box 1 cm high that grows to fit its text (18 pt lines, about 3 cm for four)."""
    box = add_box(slide, name, "\n".join(f"Line {n}" for n in range(1, lines + 1)), x=2, y=y, w=6, h=1, size=18)
    box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    return box


def test_text_box_that_grows_is_checked_at_its_grown_size(make_deck):
    import re

    from pptx.enum.shapes import MSO_SHAPE

    def fill(slide):
        growing_box(slide)
        below = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(2), Cm(3.6), Cm(6), Cm(2))
        below.name = "below"
        below.text_frame.text = "Below"

    report = run_checks(build(make_deck, fill))
    grown = finding(report, "grows-to-fit", "grows")
    assert grown["severity"] == "warning" and grown["category"] == "text"
    height = float(re.search(r"grows to ([\d.]+) cm", grown["message"]).group(1))
    assert 2.5 < height < 4.5
    overlap = finding(report, "overlap")
    assert overlap and "grows" in overlap["message"] and "below" in overlap["message"]
    assert finding(report, "overflow") is None


def test_text_box_growing_past_the_slide_is_off_slide(make_deck):
    report = run_checks(build(make_deck, lambda s: growing_box(s, y=17.5)))  # the default slide is 19.05 cm high
    assert finding(report, "off-slide", "grows")["severity"] == "error"


def test_text_box_that_fits_its_box_is_not_reported(make_deck):
    report = run_checks(build(make_deck, lambda s: growing_box(s, lines=1)))
    assert finding(report, "grows-to-fit") is None and report["summary"]["error"] == 0
