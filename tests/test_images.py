from utils import ComposedDocument
from utils import FixtureDocument


def test_images():
    doc = FixtureDocument("images.docx")
    composed = ComposedDocument(
        "master.docx", "images.docx")

    assert composed == doc
