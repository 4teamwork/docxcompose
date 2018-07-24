from docx import Document
from docxcompose.composer import Composer
import os.path


def docx_path(filename):
    return os.path.join(os.path.dirname(__file__), 'docs', filename)


class ComparableDocument(object):
    """Test helper to compare two docx documents."""

    def __init__(self, doc):
        self.doc = doc
        self.parts = self.doc.part.package.parts
        self.partnames = sorted([p.partname for p in self.parts])
        self.has_neq_partnames = False
        self.neq_parts = []

    def __eq__(self, other):
        self.has_neq_partnames = self.partnames != other.partnames
        if self.has_neq_partnames:
            return False

        for my_part, other_part in zip(self.parts, other.parts):
            if my_part.blob != other_part.blob:
                self.neq_parts.append((my_part, other_part))
        if self.neq_parts:
            return False

        return True


class ComposedComparableDocument(ComparableDocument):
    """Compose at least two documents and provide a docx document for
    comparison.

    """
    def __init__(self, master_filename, filename, *filenames):
        composer = Composer(Document(docx_path(master_filename)))
        for filename in [filename] + filenames:
            composer.append(Document(docx_path(filename)))

        super(ComposedComparableDocument, self).__init__(composer.doc)
