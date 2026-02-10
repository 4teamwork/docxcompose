from utils import ComposedDocument
from utils import FixtureDocument


def test_images():
    doc = FixtureDocument("images.docx")
    composed = ComposedDocument("master.docx", "images.docx")

    assert composed == doc


def test_embedded_and_external_image():
    doc = FixtureDocument("embedded_and_external_image.docx")
    composed = ComposedDocument("master.docx", "embedded_and_external_image.docx")

    assert composed == doc


def test_renumbering_of_non_visual_properties():
    expected = FixtureDocument("renumbering_nv_props.docx")
    composed = ComposedDocument("header_with_image.docx", "image.docx")

    assert composed == expected
