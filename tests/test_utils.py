from lxml import etree

from docxcompose.utils import xml_elements_equal


def test_xml_elements_are_equal():
    xml1 = """
    <root>
        <a x="1">Foo</a>
        <b>Bar</b>
        <ignore_me>123</ignore_me>
    </root>
    """
    xml2 = """
    <root>
        <b>Bar</b>
        <a x="1">Foo</a>
        <ignore_me>999</ignore_me>
    </root>
    """
    e1 = etree.fromstring(xml1)
    e2 = etree.fromstring(xml2)
    assert xml_elements_equal(e1, e2, ignored_tags=["ignore_me"]) is True


def test_xml_elements_are_not_equal():
    xml1 = """
    <root>
        <a x="1">Foo</a>
        <b>Bar</b>
        <ignore_me>123</ignore_me>
    </root>
    """
    xml2 = """
    <root>
        <b>Bar</b>
        <a x="2">Foo</a>
        <ignore_me>999</ignore_me>
    </root>
    """
    e1 = etree.fromstring(xml1)
    e2 = etree.fromstring(xml2)
    assert xml_elements_equal(e1, e2, ignored_tags=["ignore_me"]) is False
