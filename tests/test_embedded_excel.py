from utils import ComposedDocument
from utils import FixtureDocument


def test_hyperlinks():
    doc = FixtureDocument("embedded_excel_chart.docx")
    composed = ComposedDocument("master.docx", "embedded_excel_chart.docx")

    assert composed == doc
