from docx import Document
from docxcompose.composer import Composer
from utils import docx_path
import pytest


def test_contains_predefined_styles_in_masters_language(merged_styles):
    style_ids = [s.style_id for s in merged_styles.doc.styles]
    assert 'Heading1' in style_ids
    assert 'Heading1' in style_ids
    assert 'Strong' in style_ids
    assert 'Quote' in style_ids


def test_does_not_contain_predefined_styles_in_appended_language(merged_styles):
    style_ids = [s.style_id for s in merged_styles.doc.styles]
    assert 'berschrift1' not in style_ids
    assert 'berschrift2' not in style_ids
    assert 'Fett' not in style_ids
    assert 'Zitat' not in style_ids


def test_contains_custom_styles_from_both_docs(merged_styles):
    style_ids = [s.style_id for s in merged_styles.doc.styles]
    assert 'MyStyle1' in style_ids
    assert 'MyStyle1Char' in style_ids
    assert 'MeineFormatvorlage' in style_ids
    assert 'MeineFormatvorlageZchn' in style_ids


@pytest.fixture
def merged_styles():
    composer = Composer(Document(docx_path("styles_en.docx")))
    composer.append(Document(docx_path("styles_de.docx")))
    return composer
