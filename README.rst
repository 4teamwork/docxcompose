
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


The docxcompose console script
------------------------------


The ``docxcompose`` console script allows to compose docx files form the command
line, e.g.:

.. code:: sh

    $ docxcompose files/master.docx files/content.docx -o files/composed.docx


A note about testing
--------------------

The tests provide helpers for blackbox testing that can compare whole word
files. To do so the following files should be provided:

- a file for the expected output that should be added to the folder
  `docs/composed_fixture`
- multiple files that can be composed into the file above should be added
  to the folder `docs`.

The expected output can now be tested as follows:


.. code:: python

    def test_example():
        fixture = FixtureDocument("expected.docx")
        composed = ComposedDocument("master.docx", "slave1.docx", "slave2.docx")
        assert fixture == composed


Should the assertion fail the output file will be stored in the folder
`docs/composed_debug` with the filename of the fixture file, `expected.docx`
in case of this example.
