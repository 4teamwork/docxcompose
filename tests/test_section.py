from utils import ComposedDocument
from utils import FixtureDocument


def test_section_types_are_correct():
    composed = ComposedDocument(
        "continous_section_break.docx", "continous_section_break.docx"
    )
    assert [s.start_type for s in composed.doc.sections] == [2, 0, 0]


def test_continuous_continuous_section_break():
    doc = FixtureDocument("continous_section_break.docx")
    composed = ComposedDocument(
        "continous_section_break.docx", "continous_section_break.docx"
    )

    assert composed == doc


def test_continuous_odd_section_break():
    doc = FixtureDocument("continous_odd_section_break.docx")
    composed = ComposedDocument(
        "continous_section_break.docx", "odd_section_break.docx"
    )

    assert composed == doc


def test_next_page_section_break():
    doc = FixtureDocument("next_page_section_break.docx")
    composed = ComposedDocument("next_page_section_break.docx", "legacy_empty.docx")

    assert composed == doc


def test_next_page_section_break_in_otherwise_empty_file_preseved():
    doc = FixtureDocument("next_page_section_break_empty.docx")
    composed = ComposedDocument("master.docx", "next_page_section_break_empty.docx")

    assert composed == doc
