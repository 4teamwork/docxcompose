from docx import Document
from docxcompose.composer import Composer
from utils import docx_path
import pytest


@pytest.mark.tb
def test_contains_predefined_styles_in_masters_language(numberings_in_styles):
    import pdb; pdb.set_trace()


@pytest.fixture
def numberings_in_styles():
    composer = Composer(Document(docx_path("master.docx")))
    composer.append(Document(docx_path("numberings_styles.docx")))
    return composer
