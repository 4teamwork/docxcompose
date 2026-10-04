import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.styles.styles import Styles
from utils import ComparableDocument
from utils import ComposedDocument
from utils import docx_path
from utils import FixtureDocument

from docxcompose.composer import Composer


def test_contains_predefined_styles_in_masters_language(merged_styles):
    style_ids = [s.style_id for s in merged_styles.doc.styles]
    assert "Heading1" in style_ids
    assert "Heading1" in style_ids
    assert "Strong" in style_ids
    assert "Quote" in style_ids


def test_does_not_contain_predefined_styles_in_appended_language(merged_styles):
    style_ids = [s.style_id for s in merged_styles.doc.styles]
    assert "berschrift1" not in style_ids
    assert "berschrift2" not in style_ids
    assert "Fett" not in style_ids
    assert "Zitat" not in style_ids


def test_contains_custom_styles_from_both_docs(merged_styles):
    style_ids = [s.style_id for s in merged_styles.doc.styles]
    assert "MyStyle1" in style_ids
    assert "MyStyle1Char" in style_ids
    assert "MeineFormatvorlage" in style_ids
    assert "MeineFormatvorlageZchn" in style_ids


def test_contains_linked_styles(merged_styles):
    style_ids = [s.style_id for s in merged_styles.doc.styles]
    assert "QuoteChar" in style_ids


def test_merged_styles_de():
    doc = FixtureDocument("styles_de.docx")
    composed = ComposedDocument("styles_de.docx", "styles_en.docx")

    assert composed == doc


def test_merged_styles_en():
    doc = FixtureDocument("styles_en.docx")
    composed = ComposedDocument("styles_en.docx", "styles_de.docx")

    assert composed == doc


def test_styles_are_not_switched_for_first_numbering_element():
    doc = FixtureDocument("switched_listing_style.docx")
    composed = ComposedDocument(
        "master_switched_listing_style.docx", "switched_listing_style.docx"
    )

    assert composed == doc


def test_continue_when_no_styles():
    """Expects not to throw a type error"""
    ComposedDocument("aatmay.docx", "aatmay.docx")


def test_preserve_styles_with_same_id():
    composer = Composer(
        Document(docx_path("styles_preserve1.docx")), preserve_styles=True
    )
    composer.append(Document(docx_path("styles_preserve2.docx")))
    style_ids = [s.style_id for s in composer.doc.styles]
    assert "MyCustomStyle" in style_ids
    assert "MyCustomStyle_1" in style_ids

    expected = FixtureDocument("styles_preserve.docx")
    composed = ComparableDocument(composer.doc)
    assert composed == expected


def test_ignore_styles_with_same_id():
    composer = Composer(Document(docx_path("styles_preserve1.docx")))
    composer.append(Document(docx_path("styles_preserve2.docx")))
    style_ids = [s.style_id for s in composer.doc.styles]
    assert "MyCustomStyle" in style_ids
    assert "MyCustomStyle_1" not in style_ids


def test_preserve_styles_does_not_duplicate_identical_styles():
    composer = Composer(
        Document(docx_path("styles_preserve1.docx")), preserve_styles=True
    )
    composer.append(Document(docx_path("styles_preserve2.docx")))
    composer.append(Document(docx_path("styles_preserve2.docx")))
    composer.append(Document(docx_path("styles_preserve1.docx")))
    assert [
        s.style_id
        for s in composer.doc.styles
        if s.style_id.startswith("MyCustomStyle")
    ] == ["MyCustomStyle", "MyCustomStyleZchn", "MyCustomStyle_1"]


def test_retain_formatting_from_default_styles():
    composer = Composer(
        Document(docx_path("styles_default1.docx")), preserve_styles=True
    )
    composer.append(Document(docx_path("styles_default2.docx")))
    composed = ComparableDocument(composer.doc)
    expected = FixtureDocument("styles_default.docx")
    assert composed == expected


@pytest.mark.parametrize("preserve_styles", [False, True])
def test_style_enumerations_do_not_scale_with_paragraph_count(
    monkeypatch, preserve_styles
):
    original_iter = Styles.__iter__
    target = {"element": None, "enumerations": 0}

    def counted_iter(styles):
        if styles.element is target["element"]:
            target["enumerations"] += 1
        return original_iter(styles)

    monkeypatch.setattr(Styles, "__iter__", counted_iter)
    counts = []
    for paragraph_count in (1, 30):
        master = Document()
        source = Document()
        for _ in range(paragraph_count):
            source.add_paragraph("Repeated heading", style="Heading 1")
        target.update(element=master.styles.element, enumerations=0)
        Composer(master, preserve_styles=preserve_styles).append(source)
        counts.append(target["enumerations"])
    assert counts[0] == counts[1]


@pytest.mark.parametrize(
    "preserve_styles,same_paragraph", [(False, True), (False, False), (True, False)]
)
def test_linked_styles_are_not_duplicated(preserve_styles, same_paragraph):
    source = Document()
    paragraph_style = source.styles.add_style("Imported", WD_STYLE_TYPE.PARAGRAPH)
    character_style = source.styles.add_style("ImportedChar", WD_STYLE_TYPE.CHARACTER)
    link = OxmlElement("w:link")
    link.set(qn("w:val"), character_style.style_id)
    paragraph_style.element.append(link)
    paragraph = source.add_paragraph(style=paragraph_style)
    if not same_paragraph:
        paragraph = source.add_paragraph()
    paragraph.add_run("Linked style").style = character_style

    composer = Composer(Document(), preserve_styles=preserve_styles)
    composer.append(source)
    ids = [style.style_id for style in composer.doc.styles]
    assert ids.count("Imported") == 1
    assert ids.count("ImportedChar") == 1
    assert composer.doc.paragraphs[0].style.style_id == "Imported"
    assert composer.doc.paragraphs[-1].runs[0].style.style_id == "ImportedChar"


@pytest.mark.parametrize("preserve_styles", [False, True])
def test_styles_added_to_master_between_appends_are_seen(preserve_styles):
    composer = Composer(Document(), preserve_styles=preserve_styles)
    composer.append(Document())
    master_style = composer.doc.styles.add_style("External", WD_STYLE_TYPE.PARAGRAPH)
    master_style.font.italic = True
    source = Document()
    source_style = source.styles.add_style("External", WD_STYLE_TYPE.PARAGRAPH)
    source_style.font.bold = True
    source.add_paragraph("External style", style=source_style)

    composer.append(source)
    ids = [style.style_id for style in composer.doc.styles]
    assert ids.count("External") == 1
    assert composer.doc.styles["External"].font.italic is True
    expected_id = "External_1" if preserve_styles else "External"
    assert composer.doc.paragraphs[-1].style.style_id == expected_id


@pytest.fixture
def merged_styles():
    composer = Composer(Document(docx_path("styles_en.docx")))
    composer.append(Document(docx_path("styles_de.docx")))
    return composer
