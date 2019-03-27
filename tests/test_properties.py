from docx import Document
from docxcompose.properties import CustomProperties
from docxcompose.utils import xpath
from utils import docx_path


XPATH_CACHED_DOCPROPERTY_VALUES = 'w:r[preceding-sibling::w:r/w:fldChar/@w:fldCharType="separate"]/w:t'


def test_updates_doc_properties_with_umlauts():
    document = Document(docx_path('outdated_docproperty_with_umlauts.docx'))

    text = xpath(
        document.element.body,
        u'.//w:fldSimple[contains(@w:instr, \'DOCPROPERTY "F\xfc\xfc"\')]//w:t')
    assert u'xxx' == text[0].text

    CustomProperties(document).update_all()

    text = xpath(
        document.element.body,
        u'.//w:fldSimple[contains(@w:instr, \'DOCPROPERTY "F\xfc\xfc"\')]//w:t')
    assert u'j\xe4ja.' == text[0].text


def test_updates_doc_properties_without_quotes():
    document = Document(docx_path('outdated_docproperty_without_quotes.docx'))
    paragraphs = xpath(document.element.body, '//w:p')
    assert 1 == len(paragraphs), 'input file contains one paragraph'
    assert 1 == len(xpath(document.element.body, '//w:instrText')), \
        'input contains one complex field docproperty'
    w_p = paragraphs[0]
    cached_value = xpath(w_p, XPATH_CACHED_DOCPROPERTY_VALUES)
    assert 1 == len(cached_value), 'doc property value is in one part'
    assert 'someval' == cached_value[0].text

    CustomProperties(document).update_all()

    w_p = xpath(document.element.body, '//w:p')[0]
    cached_value = xpath(w_p, XPATH_CACHED_DOCPROPERTY_VALUES)
    assert 1 == len(cached_value), 'should have one doc property value'
    assert 'newval' == cached_value[0].text


def test_complex_docprop_fields_with_multiple_textnodes_are_updated():
    document = Document(docx_path('spellchecked_docproperty.docx'))
    paragraphs = xpath(document.element.body, '//w:p')
    assert 1 == len(paragraphs), 'input file contains one paragraph'
    assert 1 == len(xpath(document.element.body, '//w:instrText')), \
        'input contains one complex field docproperty'
    w_p = paragraphs[0]
    cached_value = xpath(w_p, XPATH_CACHED_DOCPROPERTY_VALUES)
    assert 4 == len(cached_value), \
        'doc property value is scattered over 4 parts'
    assert 'i will be spllchecked!' == ''.join(
        each.text for each in cached_value)

    CustomProperties(document).update_all()

    w_p = xpath(document.element.body, '//w:p')[0]
    cached_value = xpath(w_p, XPATH_CACHED_DOCPROPERTY_VALUES)
    assert 1 == len(cached_value), \
        'doc property value has been reset to one cached value'
    assert 'i will be spllchecked!' == cached_value[0].text


def test_complex_docprop_with_multiple_textnode_in_same_run_are_updated():
    document = Document(docx_path('two_textnodes_in_run_docproperty.docx'))
    paragraphs = xpath(document.element.body, '//w:p')
    assert 1 == len(paragraphs), 'input file contains one paragraph'
    assert 1 == len(xpath(document.element.body, '//w:instrText')), \
        'input contains one complex field docproperty'

    w_p = paragraphs[0]
    cached_value = xpath(w_p, XPATH_CACHED_DOCPROPERTY_VALUES)
    assert 2 == len(cached_value), \
        'doc property value is scattered over 2 parts'
    assert 'Hello there' == ''.join(
        each.text for each in cached_value)

    CustomProperties(document).update_all()

    w_p = xpath(document.element.body, '//w:p')[0]
    cached_value = xpath(w_p, XPATH_CACHED_DOCPROPERTY_VALUES)
    assert 1 == len(cached_value), \
        'doc property value has been reset to one cached value'
    assert 'i will be spllchecked!' == cached_value[0].text
