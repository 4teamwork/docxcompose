from utils import ComposedDocument
from utils import FixtureDocument


def test_table():
    doc = FixtureDocument("table.docx")
    composed = ComposedDocument(
        "master.docx", "table.docx")

    assert composed == doc
