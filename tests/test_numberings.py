from docx import Document
from docxcompose.composer import Composer
from docxcompose.utils import xpath
from utils import docx_path
import pytest


def test_abstractnums_from_styles_are_not_duplicated(multiple_numberings):
    anums = xpath(
        multiple_numberings.doc.part.numbering_part.element,
        './/w:abstractNum[.//w:pStyle]')
    assert len(anums) == 2


def test_restart_first_numbering(multiple_numberings):
    paragraphs = xpath(multiple_numberings.doc.element.body, './/w:p')
    assert len(xpath(paragraphs[9], './/w:numId')) == 1


@pytest.fixture
def numberings_in_styles():
    composer = Composer(Document(docx_path("master.docx")))
    composer.append(Document(docx_path("numberings_styles.docx")))
    return composer


@pytest.fixture
def multiple_numberings():
    composer = Composer(Document(docx_path("numberings_styles.docx")))
    composer.append(Document(docx_path("numberings_styles.docx")))
    return composer
