
*docxcompose* is a Python library for concatenating/appending Microsoft
Word (.docx) files.


Example usage
-------------

Append a document to another document:

.. code::

    from docxcompose.composer import Composer
    from docx import Document
    master = Document("master.docx")
    composer = Composer(master)
    doc1 = Document("doc1.docx")
    composer.append(doc1)
    composer.save("combined.docx")
