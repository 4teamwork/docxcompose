from docx import Document
from docxcompose.composer import Composer
from docxcompose.utils import xpath
from utils import docx_path
import pytest


def test_header_and_footer_refs_in_paragraph_props_get_removed(header_footer):
    refs = xpath(
        header_footer.doc.element.body,
        './/w:pPr/w:sectPr/w:headerReference|.//w:pPr/w:sectPr/w:footerReference')
    assert len(refs) == 0


@pytest.fixture
def header_footer():
    composer = Composer(Document(docx_path("master.docx")))
    composer.append(Document(docx_path("header_footer_sections.docx")))
    return composer
