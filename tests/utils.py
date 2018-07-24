from docx import Document
from docxcompose.composer import Composer
from operator import attrgetter
import os.path


def docx_path(filename):
    return os.path.join(os.path.dirname(__file__), 'docs', filename)


class ComparableDocument(object):
    """Test helper to compare two docx documents."""

    def __init__(self, doc):
        self.has_neq_partnames = False
        self.neq_parts = []

        self.doc = doc
        if not doc:
            self.parts = []
            self.partnames = []
            return

        self.parts = sorted(
            self.doc.part.package.parts, key=attrgetter('partname'))
        self.partnames = sorted(p.partname for p in self.parts)

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

    def post_compare_failed(self, other):
        """Called after a failed comparison/assert."""

        pass


class FixtureDocument(ComparableDocument):
    """Load a comparable document from the composed assets."""

    def __init__(self, composed_filename):
        self.composed_filename = composed_filename

        path = docx_path(os.path.join('composed_fixture', composed_filename))
        doc = Document(path) if os.path.isfile(path) else None

        super(FixtureDocument, self).__init__(doc)


class ComposedDocument(ComparableDocument):
    """Compose at least two documents and provide a docx document for
    comparison.

    Store output document in the `composed_debug` folder when compared to a
    document from the fixture and the assertion failed.

    """
    def __init__(self, master_filename, filename, *filenames):
        composer = Composer(Document(docx_path(master_filename)))
        for filename in (filename,) + filenames:
            composer.append(Document(docx_path(filename)))

        super(ComposedDocument, self).__init__(composer.doc)

    def post_compare_failed(self, other):
        """When comparison to a document from the fixture failed store our
        result in a debug folder.

        """
        if isinstance(other, FixtureDocument):
            path = docx_path(
                os.path.join('composed_debug', other.composed_filename))
            self.doc.save(path)
