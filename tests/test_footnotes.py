from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from utils import ComposedDocument
from utils import docx_path
from utils import FixtureDocument

from docxcompose.utils import xpath


def test_footnote():
    doc = FixtureDocument("footnote.docx")
    composed = ComposedDocument("master.docx", "footnote.docx")

    assert composed == doc


def test_footnotes_with_hyperlinks():
    doc = FixtureDocument("footnotes_with_hyperlinks.docx")
    composed = ComposedDocument("master.docx", "footnotes_with_hyperlinks.docx")

    assert composed == doc


def test_footnotes_keep_their_own_hyperlinks_and_images():
    # The second document's footnote refers to its hyperlink and its image
    # as rId1 and rId2, ids which the composed footnotes part already uses
    # for the hyperlinks of the first document.
    composed = ComposedDocument(
        "master.docx", "footnotes_with_hyperlinks.docx", "footnote_with_image.docx"
    )

    part = composed.doc.part.rels.part_with_reltype(RT.FOOTNOTES)
    footnotes = parse_xml(part.blob)
    hyperlinks = [
        part.rels[rid].target_ref for rid in xpath(footnotes, ".//w:hyperlink/@r:id")
    ]
    assert hyperlinks == [
        "https://de.wikipedia.org/wiki/Schmerz",
        "https://de.wikipedia.org/wiki/Wirksamkeit_(Recht)",
        "https://example.org/footnote",
    ]

    [rid] = xpath(footnotes, ".//a:blip/@r:embed")
    image = part.rels[rid]
    assert image.reltype == RT.IMAGE
    assert image.target_part.blob == footnote_image_blob("footnote_with_image.docx")


def footnote_image_blob(filename):
    part = Document(docx_path(filename)).part.rels.part_with_reltype(RT.FOOTNOTES)
    [rid] = xpath(parse_xml(part.blob), ".//a:blip/@r:embed")
    return part.rels[rid].target_part.blob
