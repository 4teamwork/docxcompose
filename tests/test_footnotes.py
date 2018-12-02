from utils import ComposedDocument
from utils import FixtureDocument


def test_footnote():
    doc = FixtureDocument("footnote.docx")
    composed = ComposedDocument(
        "master.docx", "footnote.docx")

    assert composed == doc


def test_footnotes_with_hyperlinks():
    doc = FixtureDocument("footnotes_with_hyperlinks.docx")
    composed = ComposedDocument(
        "master.docx", "footnotes_with_hyperlinks.docx")

    assert composed == doc
