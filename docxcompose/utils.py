from docx.oxml.xmlchemy import BaseOxmlElement
from docx.oxml.ns import nsmap

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'asvg': 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
    'cp': 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
    'o': 'urn:schemas-microsoft-com:office:office',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'v': 'urn:schemas-microsoft-com:vml',
    'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
}
nsmap.update(NS)


def xpath(element, xpath_str):
    """Performs an XPath query on the given element and returns all matching
       elements.
       Works with lxml.etree._Element and with
       docx.oxml.xmlchemy.BaseOxmlElement elements.
    """
    if isinstance(element, BaseOxmlElement):
        return element.xpath(xpath_str)
    else:
        return element.xpath(xpath_str, namespaces=NS)
