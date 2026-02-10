from utils import ComposedDocument
from utils import FixtureDocument


def test_hyperlinks():
    doc = FixtureDocument("hyperlinks.docx")
    composed = ComposedDocument("master.docx", "hyperlinks.docx")

    assert composed == doc
