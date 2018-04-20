from copy import deepcopy
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.oxml import serialize_part_xml
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.section import CT_SectPr
from docx.parts.numbering import NumberingPart
from docxcompose.image import ImageWrapper
from docxcompose.properties import CustomProperties
from docxcompose.utils import NS
from docxcompose.utils import xpath

import os.path
import random


class Composer(object):

    def __init__(self, doc):
        self.doc = doc
        self.pkg = doc.part.package

        self.restart_numbering = True

        self.reset_reference_mapping()

    def reset_reference_mapping(self):
        self.num_id_mapping = {}
        self.anum_id_mapping = {}
        self._numbering_restarted = set()

    def append(self, doc, remove_property_fields=True):
        """Append the given document."""
        index = self.append_index()
        self.insert(index, doc, remove_property_fields=remove_property_fields)

    def insert(self, index, doc, remove_property_fields=True):
        """Insert the given document at the given index."""
        self.reset_reference_mapping()

        # Remove custom property fields but keep the values
        if remove_property_fields:
            cprops = CustomProperties(doc)
            for name in cprops.dict().keys():
                cprops.remove_field(name)

        self._create_style_id_mapping(doc)

        for element in doc.element.body:
            if isinstance(element, CT_SectPr):
                continue
            element = deepcopy(element)
            self.doc.element.body.insert(index, element)
            self.add_styles(doc, element)
            self.add_numberings(doc, element)
            self.restart_first_numbering(doc, element)
            self.add_images(doc, element)
            self.add_shapes(doc, element)
            self.add_footnotes(doc, element)
            self.remove_header_and_footer_references(doc, element)
            # self.add_headers(doc, element)
            # self.add_footers(doc, element)
            self.add_hyperlinks(doc.part, self.doc.part, element)
            index += 1

        self.renumber_bookmarks()
        self.renumber_docpr_ids()

    def save(self, filename):
        self.doc.save(filename)

    def append_index(self):
        section_props = self.doc.element.body.xpath('w:sectPr')
        if section_props:
            return self.doc.element.body.index(section_props[0])
        return len(self.doc.element.body)

    def add_images(self, doc, element):
        """Add images from the given document used in the given element."""
        blips = xpath(
            element,
            '(.//w:drawing/wp:anchor|.//w:drawing/wp:inline)'
            '/a:graphic/a:graphicData//pic:pic/pic:blipFill/a:blip')
        for blip in blips:
            rid = blip.get('{%s}embed' % NS['r'])
            img_part = doc.part.rels[rid].target_part

            new_img_part = self.pkg.image_parts._get_by_sha1(img_part.sha1)
            if new_img_part is None:
                image = ImageWrapper(img_part)
                new_img_part = self.pkg.image_parts._add_image_part(image)

            new_rid = self.doc.part.relate_to(new_img_part, RT.IMAGE)
            blip.set('{%s}embed' % NS['r'], new_rid)

    def add_shapes(self, doc, element):
        shapes = xpath(element, './/v:shape/v:imagedata')
        for shape in shapes:
            rid = shape.get('{%s}id' % NS['r'])
            img_part = doc.part.rels[rid].target_part

            new_img_part = self.pkg.image_parts._get_by_sha1(img_part.sha1)
            if new_img_part is None:
                image = ImageWrapper(img_part)
                new_img_part = self.pkg.image_parts._add_image_part(image)

            new_rid = self.doc.part.relate_to(new_img_part, RT.IMAGE)
            shape.set('{%s}id' % NS['r'], new_rid)

            ole_objects = xpath(shape.getparent().getparent(), './/o:OLEObject')
            for ole_object in ole_objects:
                rid = ole_object.get('{%s}id' % NS['r'])
                ole_part = doc.part.rels[rid].target_part

                partname = self._next_ole_object_partname(ole_part.partname.ext)
                content_type = CT.OFC_OLE_OBJECT
                new_ole_part = Part(
                    partname, content_type, ole_part._blob, self.pkg)

                new_rid = self.doc.part.relate_to(new_ole_part, RT.OLE_OBJECT)
                ole_object.set('{%s}id' % NS['r'], new_rid)

    def _next_ole_object_partname(self, ext):
        def ole_object_partname(n):
            return PackURI('/word/embeddings/oleObject%d.%s' % (n, ext))
        used_numbers = [
            part.partname.idx for part in self.pkg.iter_parts()
            if part.content_type == CT.OFC_OLE_OBJECT]
        for n in range(1, len(used_numbers)+1):
            if n not in used_numbers:
                return ole_object_partname(n)
        return ole_object_partname(len(used_numbers)+1)

    def add_footnotes(self, doc, element):
        """Add footnotes from the given document used in the given element."""
        footnotes_refs = element.findall('.//w:footnoteReference', NS)

        if not footnotes_refs:
            return

        footnote_part = doc.part.rels.part_with_reltype(RT.FOOTNOTES)

        my_footnote_part = self.footnote_part()

        footnotes = parse_xml(my_footnote_part.blob)
        next_id = len(footnotes) + 1

        for ref in footnotes_refs:
            id_ = ref.get('{%s}id' % NS['w'])
            element = parse_xml(footnote_part.blob)
            footnote = deepcopy(element.find('.//w:footnote[@w:id="%s"]' % id_, NS))
            footnotes.append(footnote)
            footnote.set('{%s}id' % NS['w'], str(next_id))
            ref.set('{%s}id' % NS['w'], str(next_id))
            next_id += 1
            self.add_hyperlinks(footnote_part, my_footnote_part, element)

        my_footnote_part._blob = serialize_part_xml(footnotes)

    def footnote_part(self):
        """The footnote part of the document."""
        try:
            footnote_part = self.doc.part.rels.part_with_reltype(RT.FOOTNOTES)
        except KeyError:
            # Create a new empty footnotes part
            partname = PackURI('/word/footnotes.xml')
            content_type = CT.WML_FOOTNOTES
            xml_path = os.path.join(
                os.path.dirname(__file__), 'templates', 'footnotes.xml')
            with open(xml_path, 'rb') as f:
                xml_bytes = f.read()
            footnote_part = Part(
                partname, content_type, xml_bytes, self.doc.part.package)
            self.doc.part.relate_to(footnote_part, RT.FOOTNOTES)
        return footnote_part

    def mapped_style_id(self, style_id):
        if style_id not in self._style_id2name:
            return style_id
        return self._style_name2id.get(
                self._style_id2name[style_id], style_id)

    def _create_style_id_mapping(self, doc):
        # Style ids are language-specific, but names not (always), WTF?
        # The inserted document may have another language than the composed one.
        # Thus we map the style id using the style name.
        self._style_id2name = {s.style_id: s.name for s in doc.styles}
        self._style_name2id = {s.name: s.style_id for s in self.doc.styles}

    def add_styles(self, doc, element):
        """Add styles from the given document used in the given element."""
        our_style_ids = [s.style_id for s in self.doc.styles]
        used_style_ids = set([e.val for e in xpath(
            element, './/w:tblStyle|.//w:pStyle|.//w:rStyle')])

        for style_id in used_style_ids:
            our_style_id = self.mapped_style_id(style_id)
            if our_style_id not in our_style_ids:
                style_element = deepcopy(doc.styles.element.get_by_id(style_id))
                self.doc.styles.element.append(style_element)
                self.add_numberings(doc, style_element)
                # Also add linked styles
                linked_style_ids = xpath(style_element, './/w:link/@w:val')
                if linked_style_ids:
                    linked_style_id = linked_style_ids[0]
                    our_linked_style_id = self.mapped_style_id(linked_style_id)
                    if our_linked_style_id not in our_style_ids:
                        our_linked_style = doc.styles.element.get_by_id(
                            linked_style_id)
                        self.doc.styles.element.append(deepcopy(
                            our_linked_style))
            else:
                # Create a mapping for abstractNumIds used in existing styles
                # This is used when adding numberings to avoid having multiple
                # <w:abstractNum> elements for the same style.
                style_element = doc.styles.element.get_by_id(style_id)
                if style_element is not None:
                    num_ids = xpath(style_element, './/w:numId/@w:val')
                    if num_ids:
                        anum_ids = xpath(
                            doc.part.numbering_part.element,
                            './/w:num[@w:numId="%s"]/w:abstractNumId/@w:val' % num_ids[0])
                        if anum_ids:
                            our_style_element = self.doc.styles.element.get_by_id(our_style_id)
                            our_num_ids = xpath(our_style_element, './/w:numId/@w:val')
                            if our_num_ids:
                                numbering_part = self.numbering_part()
                                our_anum_ids = xpath(
                                    numbering_part.element,
                                    './/w:num[@w:numId="%s"]/w:abstractNumId/@w:val' % our_num_ids[0])
                                if our_anum_ids:
                                    self.anum_id_mapping[int(anum_ids[0])] = int(our_anum_ids[0])

            # Replace language-specific style id with our style id
            if our_style_id != style_id and our_style_id is not None:
                style_elements = xpath(
                    element,
                    './/w:tblStyle[@w:val="%(styleid)s"]|'
                    './/w:pStyle[@w:val="%(styleid)s"]|'
                    './/w:rStyle[@w:val="%(styleid)s"]' % dict(styleid=style_id))
                for el in style_elements:
                    el.val = our_style_id
            # Update our style ids
            our_style_ids = [s.style_id for s in self.doc.styles]

    def add_numberings(self, doc, element):
        """Add numberings from the given document used in the given element."""
        # Search for numbering references
        num_ids = set(xpath(element, './/w:numId/@w:val'))
        if not num_ids:
            return

        next_num_id, next_anum_id = self._next_numbering_ids()

        src_numbering_part = doc.part.numbering_part

        for num_id in num_ids:
            if num_id in self.num_id_mapping:
                continue

            # Find the referenced <w:num> element
            res = src_numbering_part.element.xpath(
                './/w:num[@w:numId="%s"]' % num_id)
            if not res:
                continue
            num_element = deepcopy(res[0])
            num_element.numId = next_num_id

            self.num_id_mapping[num_id] = next_num_id

            anum_id = num_element.xpath('//w:abstractNumId')[0]
            if anum_id.val not in self.anum_id_mapping:
                # Find the referenced <w:abstractNum> element
                res = src_numbering_part.element.xpath(
                    './/w:abstractNum[@w:abstractNumId="%s"]' % anum_id.val)
                if not res:
                    continue
                anum_element = deepcopy(res[0])
                self.anum_id_mapping[anum_id.val] = next_anum_id
                anum_id.val = next_anum_id
                # anum_element.abstractNumId = next_anum_id
                anum_element.set('{%s}abstractNumId' % NS['w'], str(next_anum_id))

                # Make sure we have a unique nsid so numberings restart properly
                nsid = anum_element.find('.//w:nsid', NS)
                nsid.set(
                    '{%s}val' % NS['w'],
                    "{0:0{1}X}".format(random.randint(0, 0xffffffff), 8))

                self._insert_abstract_num(anum_element)
            else:
                anum_id.val = self.anum_id_mapping[anum_id.val]

            self._insert_num(num_element)

        # Fix references
        for num_id_ref in xpath(element, './/w:numId'):
            num_id_ref.val = self.num_id_mapping.get(
                num_id_ref.val, num_id_ref.val)

    def _next_numbering_ids(self):
        numbering_part = self.numbering_part()

        # Determine next unused numId (numbering starts with 1)
        current_num_ids = [
            n.numId for n in xpath(numbering_part.element, './/w:num')]
        if current_num_ids:
            next_num_id = max(current_num_ids) + 1
        else:
            next_num_id = 1

        # Determine next unused abstractNumId (numbering starts with 0)
        current_anum_ids = [
            int(n) for n in
            xpath(numbering_part.element, './/w:abstractNum/@w:abstractNumId')]
        if current_anum_ids:
            next_anum_id = max(current_anum_ids) + 1
        else:
            next_anum_id = 0

        return next_num_id, next_anum_id

    def _insert_num(self, element):
        # Find position of last <w:num> element and insert after that
        numbering_part = self.numbering_part()
        nums = numbering_part.element.xpath('.//w:num')
        if nums:
            num_index = numbering_part.element.index(nums[-1])
            numbering_part.element.insert(num_index, element)
        else:
            numbering_part.element.append(element)

    def _insert_abstract_num(self, element):
        # Find position of first <w:num> element
        # We'll insert <w:abstractNum> before that
        numbering_part = self.numbering_part()
        nums = numbering_part.element.xpath('.//w:num')
        if nums:
            anum_index = numbering_part.element.index(nums[0])
        else:
            anum_index = 0
        numbering_part.element.insert(anum_index, element)

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

    def restart_first_numbering(self, doc, element):
        if not self.restart_numbering:
            return
        style_id = xpath(element, './/w:pStyle/@w:val')
        if not style_id:
            return
        style_id = self.mapped_style_id(style_id[0])
        if style_id in self._numbering_restarted:
            return
        style_element = self.doc.styles.element.get_by_id(style_id)
        if style_element is None:
            return
        num_id = xpath(style_element, './/w:numId/@w:val')
        if not num_id:
            return
        outline_lvl = xpath(style_element, './/w:outlineLvl')
        if outline_lvl:
            # Styles with an outline level are propably headings.
            # Do not restart numbering of headings
            return

        numbering_part = self.numbering_part()
        num_element = xpath(
            numbering_part.element,
            './/w:num[@w:numId="%s"]' % num_id[0])
        anum_id = xpath(num_element[0], './/w:abstractNumId/@w:val')[0]
        anum_element = xpath(
            numbering_part.element,
            './/w:abstractNum[@w:abstractNumId="%s"]' % anum_id)
        num_fmt = xpath(
            anum_element[0], './/w:lvl[@w:ilvl="0"]/w:numFmt/@w:val')
        # Do not restart numbering of bullets
        if num_fmt[0] == 'bullet':
            return

        new_num_element = deepcopy(num_element[0])
        lvl_override = parse_xml(
            '<w:lvlOverride xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>')
        new_num_element.append(lvl_override)
        next_num_id, next_anum_id = self._next_numbering_ids()
        new_num_element.numId = next_num_id
        self._insert_num(new_num_element)

        paragraph_props = xpath(element, './/w:pPr/w:pStyle[@w:val="%s"]/parent::w:pPr' % style_id)
        num_pr = xpath(paragraph_props[0], './/w:numPr')
        if num_pr:
            num_pr = num_pr[0]
            num_pr.numId.val = next_num_id
        else:
            num_pr = parse_xml(
                '<w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:ilvl w:val="0"/><w:numId w:val="%s"/></w:numPr>' % next_num_id)
            paragraph_props[0].append(num_pr)
        self._numbering_restarted.add(style_id)

    def add_hyperlinks(self, src_part, dst_part, element):
        """Add hyperlinks from src_part referenced in element to dst_part."""
        hyperlink_refs = xpath(element, './/w:hyperlink')
        for hyperlink_ref in hyperlink_refs:
            rid = hyperlink_ref.get('{%s}id' % NS['r'])
            if rid is None:
                continue
            rel = src_part.rels[rid]
            if rel.is_external:
                new_rid = dst_part.rels.get_or_add_ext_rel(
                    rel.reltype, rel.target_ref)
                hyperlink_ref.set('{%s}id' % NS['r'], new_rid)

    def add_headers(self, doc, element):
        header_refs = xpath(element, './/w:headerReference')
        if not header_refs:
            return
        for ref in header_refs:
            rid = ref.get('{%s}id' % NS['r'])
            rel = doc.part.rels[rid]
            header_part = self.header_part(content=rel.target_part.blob)
            my_rel = self.doc.part.rels.get_or_add(
                rel.reltype, header_part)
            ref.set('{%s}id' % NS['r'], my_rel.rId)

    def header_part(self, content=None):
        """The header part of the document."""
        header_rels = [
            rel for rel in self.doc.part.rels.values() if rel.reltype == RT.HEADER]
        next_id = len(header_rels) + 1
        # Create a new header part
        partname = PackURI('/word/header%s.xml' % next_id)
        content_type = CT.WML_HEADER
        if not content:
            xml_path = os.path.join(
                os.path.dirname(__file__), 'templates', 'header.xml')
            with open(xml_path, 'rb') as f:
                content = f.read()
        header_part = Part(
            partname, content_type, content, self.doc.part.package)
        self.doc.part.relate_to(header_part, RT.HEADER)
        return header_part

    def add_footers(self, doc, element):
        footer_refs = xpath(element, './/w:footerReference')
        if not footer_refs:
            return
        for ref in footer_refs:
            rid = ref.get('{%s}id' % NS['r'])
            rel = doc.part.rels[rid]
            footer_part = self.footer_part(content=rel.target_part.blob)
            my_rel = self.doc.part.rels.get_or_add(
                rel.reltype, footer_part)
            ref.set('{%s}id' % NS['r'], my_rel.rId)

    def footer_part(self, content=None):
        """The footer part of the document."""
        footer_rels = [
            rel for rel in self.doc.part.rels.values() if rel.reltype == RT.FOOTER]
        next_id = len(footer_rels) + 1
        # Create a new header part
        partname = PackURI('/word/footer%s.xml' % next_id)
        content_type = CT.WML_FOOTER
        if not content:
            xml_path = os.path.join(
                os.path.dirname(__file__), 'templates', 'footer.xml')
            with open(xml_path, 'rb') as f:
                content = f.read()
        footer_part = Part(
            partname, content_type, content, self.doc.part.package)
        self.doc.part.relate_to(footer_part, RT.FOOTER)
        return footer_part

    def remove_header_and_footer_references(self, doc, element):
        refs = xpath(
            element, './/w:headerReference|.//w:footerReference')
        for ref in refs:
            ref.getparent().remove(ref)

    def renumber_bookmarks(self):
        bookmarks_start = xpath(self.doc.element.body, './/w:bookmarkStart')
        bookmark_id = 0
        for bookmark in bookmarks_start:
            bookmark.set('{%s}id' % NS['w'], str(bookmark_id))
            bookmark_id += 1
        bookmarks_end = xpath(self.doc.element.body, './/w:bookmarkEnd')
        bookmark_id = 0
        for bookmark in bookmarks_end:
            bookmark.set('{%s}id' % NS['w'], str(bookmark_id))
            bookmark_id += 1

    def renumber_docpr_ids(self):
        doc_prs = xpath(
            self.doc.element.body, './/wp:docPr')
        doc_pr_id = 1
        for doc_pr in doc_prs:
            doc_pr.id = doc_pr_id
            doc_pr_id += 1
