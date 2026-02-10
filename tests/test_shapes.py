from utils import ComposedDocument
from utils import FixtureDocument


def test_shapes():
    doc = FixtureDocument("embedded_visio.docx")
    composed = ComposedDocument("master.docx", "embedded_visio.docx")

    assert composed == doc
