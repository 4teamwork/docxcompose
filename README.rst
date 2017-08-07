
*docxmerge* is a Python library for concatenating/appending Microsoft
Word (.docx) files.


Example usage
-------------

Append a document to another document:

.. code::

    from docxmerge.builder import DocumentBuilder
    from docx import Document
    master = Document("master.docx")
    builder = DocumentBuilder(master)
    doc1 = Document("doc1.docx")
    builder.append(doc1)
    builder.save("combined.docx")
