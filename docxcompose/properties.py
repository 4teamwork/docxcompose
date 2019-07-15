from copy import deepcopy
from datetime import datetime
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.oxml import serialize_part_xml
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.coreprops import CT_CoreProperties
from docxcompose.utils import NS
from docxcompose.utils import xpath
from lxml.etree import QName
from six import string_types
from six import text_type
import pkg_resources
import re


CUSTOM_PROPERTY_FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'
CUSTOM_PROPERTY_TYPES = {
    'text': '<vt:lpwstr xmlns:vt="{}"/>'.format(NS['vt']),
    'int': '<vt:i4 xmlns:vt="{}"/>'.format(NS['vt']),
    'bool': '<vt:bool xmlns:vt="{}"/>'.format(NS['vt']),
    'datetime': '<vt:filetime xmlns:vt="{}"/>'.format(NS['vt']),
    'float': '<vt:r8 xmlns:vt="{}"/>'.format(NS['vt']),
}
MIN_PID = 2  # Property IDs have to start with 2


def value2vt(value):
    if isinstance(value, bool):
        el = parse_xml(CUSTOM_PROPERTY_TYPES['bool'])
        el.text = 'true' if value else 'false'
    elif isinstance(value, int):
        el = parse_xml(CUSTOM_PROPERTY_TYPES['int'])
        el.text = text_type(value)
    elif isinstance(value, float):
        el = parse_xml(CUSTOM_PROPERTY_TYPES['float'])
        el.text = text_type(value)
    elif isinstance(value, datetime):
        el = parse_xml(CUSTOM_PROPERTY_TYPES['datetime'])
        el.text = value.strftime('%Y-%m-%dT%H:%M:%SZ')
    elif isinstance(value, string_types):
        el = parse_xml(CUSTOM_PROPERTY_TYPES['text'])
        el.text = value
    else:
        raise TypeError('Unsupported type {}'.format(type(value)))
    return el


def vt2value(element):
    tag = QName(element).localname
    if tag == 'bool':
        if element.text.lower() == u'true':
            return True
        else:
            return False
    elif tag in ['i1', 'i2', 'i4', 'int', 'ui1', 'ui2', 'ui4', 'uint']:
        return int(element.text)
    elif tag in ['r4', 'r8']:
        return float(element.text)
    elif tag == 'filetime':
        return CT_CoreProperties._parse_W3CDTF_to_datetime(element.text)
    else:
        return element.text


class CustomProperties(object):
    """Custom doc properties stored in ``/docProps/custom.xml``.
       Allows updating of doc properties in a document.
    """
    def __init__(self, doc):
        self.doc = doc
        self.part = None
        self._element = None

        try:
            part = doc.part.package.part_related_by(RT.CUSTOM_PROPERTIES)
        except KeyError:
            self._element = parse_xml(self._part_template())
        else:
            self.part = part
            self._element = parse_xml(part.blob)

    def _part_template(self):
        return pkg_resources.resource_string(
            'docxcompose', 'templates/custom.xml')

    def _update_part(self):
        if self.part is None:
            # Create a new part for custom properties
            partname = PackURI('/docProps/custom.xml')
            self.part = Part(
                partname, CT.OFC_CUSTOM_PROPERTIES,
                serialize_part_xml(self._element), self.doc.part.package)
            self.doc.part.package.relate_to(self.part, RT.CUSTOM_PROPERTIES)
            self._element = parse_xml(self.part.blob)
        else:
            self.part._blob = serialize_part_xml(self._element)

    def __getitem__(self, key):
        """Get the value of a property."""
        props = xpath(
            self._element,
            u'.//cp:property[@name="{}"]'.format(key))

        if not props:
            raise KeyError(key)

        return vt2value(props[0][0])

    def __setitem__(self, key, value):
        """Set the value of a property."""
        props = xpath(
            self._element,
            u'.//cp:property[@name="{}"]'.format(key))
        if not props:
            self.add(key, value)
            return

        value_el = props[0][0]
        new_value_el = value2vt(value)
        value_el.getparent().replace(value_el, new_value_el)

        self._update_part()

    def __delitem__(self, key):
        """Delete a property."""
        props = xpath(
            self._element,
            u'.//cp:property[@name="{}"]'.format(key))

        if not props:
            raise KeyError(key)

        props[0].getparent().remove(props[0])
        # Renumber pids
        pid = MIN_PID
        for prop in self._element:
            prop.set('pid', text_type(pid))
            pid += 1

        self._update_part()

    def __contains__(self, item):
        props = xpath(
            self._element,
            u'.//cp:property[@name="{}"]'.format(item))
        if props:
            return True
        else:
            return False

    def get(self, key, default=None):
        try:
            return self[key]
        except KeyError:
            return default

    def add(self, name, value):
        """Add a property."""
        pids = [int(pid) for pid in xpath(self._element, u'.//cp:property/@pid')]
        if pids:
            pid = max(pids) + 1
        else:
            pid = MIN_PID
        prop = parse_xml('<cp:property xmlns:cp="{}"/>'.format(NS['cp']))
        prop.set('fmtid', CUSTOM_PROPERTY_FMTID)
        prop.set('name', name)
        prop.set('pid', text_type(pid))
        value_el = value2vt(value)
        prop.append(value_el)
        self._element.append(prop)

        self._update_part()

    def keys(self):
        if self._element is None:
            return []

        props = xpath(self._element, u'.//cp:property')
        return [prop.get('name') for prop in props]

    def values(self):
        if self._element is None:
            return []

        props = xpath(self._element, u'.//cp:property')
        return [vt2value(prop[0]) for prop in props]

    def items(self):
        if self._element is None:
            return []

        props = xpath(self._element, u'.//cp:property')
        return [(prop.get('name'), vt2value(prop[0])) for prop in props]

    def set_properties(self, properties):
        for name, value in properties.items():
            self.set(name, value)

    def find_docprops_in_document(self, name=None):
        """This method searches for all doc-properties in the document
        """
        # First we search for the simple fields:
        sfield_nodes = xpath(
            self.doc.element.body,
            u'.//w:fldSimple[contains(@w:instr, \'DOCPROPERTY \')]')

        docprops = [SimpleField(sfield_node) for sfield_node in sfield_nodes]

        # Now for the complex fields
        cfield_nodes = xpath(
            self.doc.element.body,
            u'.//w:instrText[contains(.,\'DOCPROPERTY \')]')

        docprops.extend([ComplexField(cfield_node) for cfield_node in cfield_nodes])

        if name is not None:
            docprops = filter(lambda prop: prop.name == name, docprops)
        return docprops

    def update_all(self):
        """Update all the document's doc-properties."""
        docprops = self.find_docprops_in_document()
        available_docprops = dict(self.items())

        for docprop in docprops:
            value = available_docprops.get(docprop.name)
            if value is None:
                continue
            docprop.update(value)

    def update(self, name, value):
        """Update all instances of a given doc-property in the document."""
        docprops = self.find_docprops_in_document(name)
        for docprop in docprops:
            docprop.update(value)

    def remove_field(self, name):
        """Remove the property field but keep it's value."""

        # Simple field
        sfield = xpath(
            self.doc.element.body,
            u'.//w:fldSimple[contains(@w:instr, \'DOCPROPERTY "{}"\')]'.format(name))
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
            u'.//w:instrText[contains(.,\'DOCPROPERTY "{}"\')]'.format(name))
        if cfield:
            w_p = cfield[0].getparent().getparent()
            # Create list of <w:r> nodes for removal
            # Get all <w:r> nodes between <w:fldChar w:fldCharType="begin"/>
            # and <w:fldChar w:fldCharType="separate"/> including boundaries.
            w_rs = xpath(
                w_p,
                u'.//w:r[following-sibling::w:r/w:fldChar/@w:fldCharType="separate" '
                u'and preceding-sibling::w:r/w:fldChar/@w:fldCharType="begin" '
                u'or self::w:r/w:fldChar/@w:fldCharType="begin" '
                u'or self::w:r/w:fldChar/@w:fldCharType="separate"]')
            # Also include <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            w_rs.extend(xpath(
                w_p, u'.//w:r/w:fldChar[@w:fldCharType="end"]/parent::w:r'))
            for w_r in w_rs:
                w_p.remove(w_r)


class FieldBase(object):

    fieldname_search_expr = re.compile(
            r'DOCPROPERTY +"{0,1}([^\\]*?)"{0,1} +\\\* MERGEFORMAT',
            flags=re.UNICODE)

    def __init__(self, field_node):
        self.node = field_node
        self.name = self._get_fieldname()

    def _format_value(self, value):
        if isinstance(value, bool):
            return u'Y' if value else u'N'
        elif isinstance(value, datetime):
            return value.strftime('%x')
        else:
            return text_type(value)

    def update(self, value):
        raise NotImplementedError()

    def _get_fieldname_string(self):
        raise NotImplementedError()

    def _get_fieldname(self):
        match = self.fieldname_search_expr.search(self._get_fieldname_string())
        if match is None:
            return None
        return match.groups()[0]


class RunNode(object):

    def __init__(self, node, is_begin=False, is_end=False,
                 is_separate=False, is_after_separate=False):
        self.node = node
        self.is_begin = is_begin
        self.is_end = is_end
        self.is_separate = is_separate
        self.is_after_separate = is_after_separate


class SimpleField(FieldBase):

    attr_name = "{{{}}}instr".format(NS["w"])

    def _get_fieldname_string(self):
        return self.node.attrib[self.attr_name]

    def update(self, value):
        text = xpath(self.node, './/w:t')
        if text:
            text[0].text = self._format_value(value)


class ComplexField(FieldBase):

    def __init__(self, field_node):
        super(ComplexField, self).__init__(field_node)
        # run and paragraph containing the field
        self.w_r = self.node.getparent()
        self.w_p = self.w_r.getparent()

    def _get_fieldname_string(self):
        return self.node.text

    def _get_field_bound_indexes(self):
        """Searches for the begin and end tags surrounding the complex
        field as well as the corresponding position of the separate tag.
        Returns their position in the parent node"""
        fieldname_index = self.w_p.index(self.w_r)
        begins = [self.w_p.index(el.getparent()) for el in xpath(self.w_p, u'.//w:r/w:fldChar[@w:fldCharType="begin"]')]
        separates = [self.w_p.index(el.getparent()) for el in xpath(self.w_p, u'.//w:r/w:fldChar[@w:fldCharType="separate"]')]
        ends = [self.w_p.index(el.getparent()) for el in xpath(self.w_p, u'.//w:r/w:fldChar[@w:fldCharType="end"]')]
        for begin, separate, end in zip(begins, separates, ends):
            if begin < fieldname_index < end:
                return begin, separate, end
        return None

    def get_runs(self):
        """Get run fields from the current complex field"""
        runs = []
        bounds = self._get_field_bound_indexes()
        if bounds is None:
            return runs
        begin, separate, end = bounds

        for run in xpath(self.w_p, u'.//w:r'):
            position = self.w_p.index(run)
            if begin <= position <= end:
                runs.append(RunNode(run, position == begin, position == end,
                                    position == separate, position > separate))
        return runs

    def update(self, value):
        runs = self.get_runs()

        # only keep run fields after <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        runs_after_separate = [run for run in runs
                               if run.is_after_separate and not run.is_end]

        if runs_after_separate:
            first_w_r = runs_after_separate[0].node
            text = xpath(first_w_r, u'.//w:t')
            if text:
                text[0].text = self._format_value(value)
            # remove any additional text-nodes inside the first run. we
            # update the first text-node only with the full cached
            # docproperty value. if for some reason the initial cached
            # value is split into multiple text nodes we remove any
            # additional node after updating the first node.
            for unnecessary_w_t in text[1:]:
                first_w_r.remove(unnecessary_w_t)

            # if there are multiple runs between "separate" and "end" they
            # all may contain a piece of the cached docproperty value. we
            # can't reliably handle this situation and only update the
            # first node in the first run with the full cached value. it
            # appears any additional runs with text nodes should then be
            # removed to avoid duplicating parts of the cached docproperty
            # value.
            for run in runs_after_separate[1:]:
                text = xpath(run.node, u'.//w:t')
                if text:
                    self.w_p.remove(run.node)
