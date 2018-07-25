from utils import ComposedDocument
from utils import FixtureDocument


def test_url():
    doc = FixtureDocument("url.docx")
    composed = ComposedDocument(
        "master.docx", "url.docx")

    assert composed == doc
