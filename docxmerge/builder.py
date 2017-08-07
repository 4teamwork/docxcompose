from copy import deepcopy
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.oxml import serialize_part_xml
from docx.opc.packuri import PackURI
from docx.oxml import parse_xml
from docx.parts.numbering import NumberingPart
import os.path

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


class DocumentBuilder(object):

    def __init__(self, doc):
        self.doc = doc
        self.pkg = doc.part.package

        self.reset_reference_mapping()

    def reset_reference_mapping(self):
        self.num_id_mapping = {}

    def append(self, doc):
        """Append the given document."""
        self.reset_reference_mapping()
        for element in doc.element.body:
            element = deepcopy(element)
            self.doc.element.body.append(element)
            self.add_styles(doc, element)
            self.add_numberings(doc, element)
            self.add_images(doc, element)
            self.add_footnotes(doc, element)

    def insert(self, index, doc):
        """Insert the given document at the given index."""
        self.reset_reference_mapping()
        for element in doc.element.body:
            element = deepcopy(element)
            self.doc.element.body.insert(index, element)
            self.add_styles(doc, element)
            self.add_numberings(doc, element)
            self.add_images(doc, element)
            self.add_footnotes(doc, element)
            index += 1

    def add_images(self, doc, element):
        """Add images from the given document used in the given element."""
        blips = element.findall(
            './/w:drawing/wp:inline/a:graphic/a:graphicData'
            '/pic:pic/pic:blipFill/a:blip', NS)
        for blip in blips:
            rid = blip.get('{%s}embed' % NS['r'])
            img_part = doc.part.rels[rid].target_part
            self.pkg.image_parts.append(img_part)
            new_rid = self.doc.part.relate_to(img_part, RT.IMAGE)
            blip.set('{%s}embed' % NS['r'], new_rid)

    def add_footnotes(self, doc, element):
        """Add footnotes from the given document used in the given element."""
        footnotes_refs = element.findall('.//w:footnoteReference', NS)

        if not footnotes_refs:
            return

        footnote_part = doc.part.rels.part_with_reltype(RT.FOOTNOTES)
        try:
            doc_footnote_part = self.doc.part.rels.part_with_reltype(RT.FOOTNOTES)
        except KeyError:
            doc_footnote_part = deepcopy(footnote_part)
            self.doc.part.relate_to(doc_footnote_part, RT.FOOTNOTES)
            footnotes = parse_xml(doc_footnote_part.blob)
            for footnote in footnotes[2:]:
                footnotes.remove(footnote)
            doc_footnote_part._blob = serialize_part_xml(footnotes)

        footnotes = parse_xml(doc_footnote_part.blob)
        next_id = len(footnotes) - 1

        for ref in footnotes_refs:
            id_ = ref.get('{%s}id' % NS['w'])
            element = parse_xml(footnote_part.blob)
            footnote = deepcopy(element.find('.//w:footnote[@w:id="%s"]' % id_, NS))
            footnotes.append(footnote)
            footnote.set('{%s}id' % NS['w'], str(next_id))
            next_id += 1

        doc_footnote_part._blob = serialize_part_xml(footnotes)

    def add_styles(self, doc, element):
        """Add styles from the given document used in the given element."""
        our_style_ids = [s.style_id for s in self.doc.styles]
        used_style_ids = [e.val for e in element.xpath(
            './/w:tblStyle|.//w:pStyle')]
        for style_id in used_style_ids:
            if style_id not in our_style_ids:
                style_element = doc.styles.element.get_by_id(style_id)
                self.doc.styles.element.append(deepcopy(style_element))

    def add_numberings(self, doc, element):
        """Add numberings from the given document used in the given element."""

        # Search for numbering references
        num_ids = set([n.val for n in element.xpath('.//w:numId')])
        if not num_ids:
            return

        numbering_part = self.numbering_part()

        # Determine next unused numId (numbering starts with 1)
        current_num_ids = [
            n.numId for n in numbering_part.element.xpath('.//w:num')]
        if current_num_ids:
            next_num_id = max(current_num_ids) + 1
        else:
            next_num_id = 1

        # Determine next unused abstractNumId (numbering starts with 0)
        current_anum_ids = [int(n.get('{%s}abstractNumId' % NS['w'])) for n in
                            numbering_part.element.xpath('.//w:abstractNum')]
        if current_anum_ids:
            next_anum_id = max(current_anum_ids) + 1
        else:
            next_anum_id = 0

        src_numbering_part = doc.part.numbering_part
        for num_id in num_ids:
            if num_id in self.num_id_mapping:
                continue
            num_element = deepcopy(src_numbering_part.element.xpath(
                './/w:num[@w:numId="%s"]' % num_id)[0])
            anum_id = num_element.xpath('//w:abstractNumId')[0]
            anum_element = deepcopy(src_numbering_part.element.xpath(
                './/w:abstractNum[@w:abstractNumId="%s"]' % anum_id.val)[0])

            self.num_id_mapping[num_id] = next_num_id

            num_element.numId = next_num_id
            anum_id.val = next_anum_id
            # anum_element.abstractNumId = next_anum_id
            anum_element.set('{%s}abstractNumId' % NS['w'], str(next_anum_id))

            # Find position of first <w:num> element
            nums = numbering_part.element.xpath('.//w:num')
            if nums:
                anum_index = numbering_part.element.index(nums[0])
            else:
                anum_index = 0

            # Insert <w:abstractNum> before <w:num> elements
            numbering_part.element.insert(anum_index, anum_element)
            numbering_part.element.append(num_element)

        # Fix references
        for num_id_ref in element.xpath('.//w:numId'):
            num_id_ref.val = self.num_id_mapping[num_id_ref.val]

    def numbering_part(self):
        """The numbering part of the document."""
        try:
            numbering_part = self.doc.part.rels.part_with_reltype(RT.NUMBERING)
        except KeyError:
            # Create a new empty numbering part
            partname = PackURI('/word/numbering.xml')
            content_type = CT.WML_NUMBERING
            xml_path = os.path.join(
                os.path.dirname(__file__), 'templates', 'numbering.xml')
            with open(xml_path, 'rb') as f:
                xml_bytes = f.read()
            element = parse_xml(xml_bytes)
            numbering_part = NumberingPart(
                partname, content_type, element, self.doc.part.package)
            self.doc.part.relate_to(numbering_part, RT.NUMBERING)
        return numbering_part
