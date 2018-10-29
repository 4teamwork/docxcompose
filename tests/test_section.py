from utils import ComposedDocument
from utils import FixtureDocument


def test_continuous_continuous_section_break():
    doc = FixtureDocument("continous_section_break.docx")
    composed = ComposedDocument(
        "continous_section_break.docx", "continous_section_break.docx")

    assert composed == doc


def test_continuous_odd_section_break():
    doc = FixtureDocument("continous_odd_section_break.docx")
    composed = ComposedDocument(
        "continous_section_break.docx", "odd_section_break.docx")

    assert composed == doc
