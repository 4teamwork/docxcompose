from __future__ import print_function
from copy import deepcopy
from datetime import datetime
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.coreprops import CT_CoreProperties
from docxcompose.utils import xpath


class CustomProperties(object):
    """Custom doc properties stored in ``/docProps/custom.xml``.
       Allows updating of doc properties in a document.
    """
    def __init__(self, doc):
        self.doc = doc
        self._element = None

        try:
            part = doc.part.package.part_related_by(RT.CUSTOM_PROPERTIES)
        except KeyError:
            pass
        else:
            self._element = parse_xml(part.blob)

    def dict(self):
        """Returns a dict with all custom doc properties"""
        if not self._element:
            return dict()

        props = xpath(self._element, './/cp:property')
        return {prop.get('name'): prop[0].text for prop in props}

    def get(self, name):
        """Get the value of a property."""
        prop = xpath(
            self._element,
            './/cp:property[@name="{}"]'.format(name))
        if prop:
            value = list(prop[0])[0]
            if value.tag.endswith('}lpwstr'):
                return value.text
            elif value.tag.endswith('}i4'):
                return int(value.text)
            elif value.tag.endswith('}bool'):
                if value.text.lower() == 'true':
                    return True
                else:
                    return False
            elif value.tag.endswith('}filetime'):
                return CT_CoreProperties._parse_W3CDTF_to_datetime(value.text)

    def update_all(self):
        """Update all the document's doc-properties."""
        for name in self.dict().keys():
            self.update(name)

    def update(self, name):
        """Update a property field value."""
        value = self.get(name)
        if isinstance(value, bool):
            value = 'Y' if value else 'N'
        elif isinstance(value, datetime):
            value = value.strftime('%x')
        else:
            value = unicode(value)

        # Simple field
        sfield = xpath(
            self.doc.element.body,
            './/w:fldSimple[contains(@w:instr, \'DOCPROPERTY "{}"\')]'.format(name))
        if sfield:
            text = xpath(sfield[0], './/w:t')
            if text:
                text[0].text = value

        # Complex field
        cfield = xpath(
            self.doc.element.body,
            './/w:instrText[contains(.,\'DOCPROPERTY "{}"\')]'.format(name))
        if cfield:
            runs = xpath(
                cfield[0].getparent().getparent(),
                './/w:r[following-sibling::w:r/w:fldChar/@w:fldCharType="end"'
                ' and preceding-sibling::w:r/w:fldChar/@w:fldCharType="separate"]')
            if runs:
                text = xpath(runs[0], './/w:t')
                if text:
                    text[0].text = value
                if len(text) > 1:
                    print("Warning! Multiple Textnodes")

    def remove_field(self, name):
        """Remove the property field but keep it's value."""

        # Simple field
        sfield = xpath(
            self.doc.element.body,
            './/w:fldSimple[contains(@w:instr, \'DOCPROPERTY "{}"\')]'.format(name))
        if sfield:
            sfield = sfield[0]
            parent = sfield.getparent()
            index = list(parent).index(sfield)
            w_r = deepcopy(sfield[0])
            parent.remove(sfield)
            parent.insert(index, w_r)

        # Complex field
        cfield = xpath(
            self.doc.element.body,
            './/w:instrText[contains(.,\'DOCPROPERTY "{}"\')]'.format(name))
        if cfield:
            w_p = cfield[0].getparent().getparent()
            # Create list of <w:r> nodes for removal
            # Get all <w:r> nodes between <w:fldChar w:fldCharType="begin"/>
            # and <w:fldChar w:fldCharType="separate"/> including boundaries.
            w_rs = xpath(
                w_p,
                './/w:r[following-sibling::w:r/w:fldChar/@w:fldCharType="separate" '
                'and preceding-sibling::w:r/w:fldChar/@w:fldCharType="begin" '
                'or self::w:r/w:fldChar/@w:fldCharType="begin" '
                'or self::w:r/w:fldChar/@w:fldCharType="separate"]')
            # Also include <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            w_rs.extend(xpath(
                w_p, './/w:r/w:fldChar[@w:fldCharType="end"]/parent::w:r'))
            for w_r in w_rs:
                w_p.remove(w_r)
