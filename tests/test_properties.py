from docx import Document
from docxcompose.properties import CustomProperties
from docxcompose.utils import xpath
from utils import docx_path


def test_updates_doc_properties_with_umlauts():
    document = Document(docx_path("outdated_docproperty_with_umlauts.docx"))

    text = xpath(
        document.element.body,
        u'.//w:fldSimple[contains(@w:instr, \'DOCPROPERTY "F\xfc\xfc"\')]//w:t')
    assert u'xxx' == text[0].text

    CustomProperties(document).update_all()

    text = xpath(
        document.element.body,
        u'.//w:fldSimple[contains(@w:instr, \'DOCPROPERTY "F\xfc\xfc"\')]//w:t')
    assert u'j\xe4ja.' == text[0].text
