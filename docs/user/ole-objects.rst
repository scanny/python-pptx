Embedded OLE objects
====================

A slide may embed an external "file" as an OLE (Object Linking and Embedding)
object -- a graphic frame on the slide that, when double-clicked in
PowerPoint, opens the file using whichever application the operating system
associates with the file's `progId`. The file bytes are stored verbatim as an
additional part inside the `.pptx` package.

Embedding a Microsoft Office file
---------------------------------

For the common Microsoft Office file types, a member of the
|PROG_ID| enumeration selects the right `progId` and part-name conventions::

    >>> from pptx.enum.shapes import PROG_ID
    >>> shape = slide.shapes.add_ole_object(
    ...     "budget.xlsx",
    ...     PROG_ID.XLSX,
    ...     left=Inches(1),
    ...     top=Inches(1),
    ... )
    >>> shape.ole_format.prog_id
    'Excel.Sheet.12'

The built-in members for Office file types are ``PROG_ID.DOCX``,
``PROG_ID.PPTX``, and ``PROG_ID.XLSX``.

Embedding an arbitrary file
---------------------------

Starting with the resolution of `issue #752`_, `add_ole_object()` accepts any
`progId` string -- not only the Office shortcuts -- together with an explicit
file-extension hint. This lets you embed a zip archive, a PDF, an HTML
document, or any other file type::

    >>> shape = slide.shapes.add_ole_object(
    ...     "archive.zip",
    ...     prog_id="Package",
    ...     left=Inches(1),
    ...     top=Inches(1),
    ...     extension="zip",
    ... )

When `object_file` is a string path, the extension is inferred from the path
name, so the `extension` keyword is optional in that case::

    >>> shape = slide.shapes.add_ole_object(
    ...     "report.pdf", prog_id="AcroExch.Document.DC", left=..., top=...
    ... )

When `object_file` is a file-like object (e.g. `BytesIO`), supply the
`extension` keyword so the embedded part receives a meaningful extension
inside the package (for example ``/ppt/embeddings/oleObject1.zip``). If no
extension can be resolved, the part defaults to the ``.bin`` extension.

For convenience, the |PROG_ID| enumeration also ships with members for the
most common generic embeds: ``PROG_ID.ZIP``, ``PROG_ID.PDF``,
``PROG_ID.DOC``, and ``PROG_ID.HTML``. Using these members sets the `progId`
and extension together::

    >>> shape = slide.shapes.add_ole_object(
    ...     "archive.zip", PROG_ID.ZIP, left=Inches(1), top=Inches(1)
    ... )
    >>> shape.ole_format.prog_id
    'Package'

Notes
-----

* The progId determines which application PowerPoint will launch when the
  icon is double-clicked. It must match an application registered on the
  target machine; PowerPoint ignores double-click when the progId is not
  registered.
* Support for particular embedded types is OS-dependent. Excel (XLSX) and
  Word (DOCX) are the most broadly supported.
* `python-pptx` only supports the "display as icon" embed mode; a preview
  image cannot be generated because there is no OLE server available in a
  pure-Python environment.

.. |PROG_ID| replace:: :class:`~pptx.enum.shapes.PROG_ID`

.. _`issue #752`:
   https://github.com/scanny/python-pptx/issues/752
