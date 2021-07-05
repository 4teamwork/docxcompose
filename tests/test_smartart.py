from utils import ComposedDocument
from utils import FixtureDocument


def test_smartart():
    doc = FixtureDocument("smart_art.docx")
    composed = ComposedDocument(
        "master.docx", "smart_art.docx")

    assert composed == doc
