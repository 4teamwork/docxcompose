from docx import Document
from docxcompose.sdt import StructuredDocumentTags
from utils import ComparableDocument
from utils import docx_path
from utils import FixtureDocument


def test_set_sdt_text_content():
    doc = Document(docx_path('content_controls.docx'))
    sdt = StructuredDocumentTags(doc)
    sdt.set_text('cc.plain_text', u'Foo Bar')
    sdt.set_text('cc.plain_text_multiline', u'Foo Bar')
    sdt.set_text('cc.plain_text_empty', u'Foo Bar')
    sdt.set_text('cc.rich_text', u'Foo Bar')

    updated = ComparableDocument(doc)
    expected = FixtureDocument("content_controls.docx")

    assert updated == expected
