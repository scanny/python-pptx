.. :changelog:

Release History
---------------

0.3.0 (2013-12-12)
++++++++++++++++++

- Add read/write font color property supporting RGB, theme color, and inherit
  color types
- Add font typeface and italic support
- Add text frame margins and word-wrap
- Add support for external relationships, e.g. linked spreadsheet
- Add hyperlink support for text run in shape and table cell
- Add fill color and brightness for shape and table cell, fill can also be set
  to transparent (no fill)
- Add read/write position and size properties to shape and picture
- Replace PIL dependency with Pillow
- Restructure modules to better suit size of library


0.2.6 (2013-06-22)
++++++++++++++++++

- Add read/write access to core document properties
- Hotfix to accomodate connector shapes in _AutoShapeType
- Hotfix to allow customXml parts to load when present


0.2.5 (2013-06-11)
++++++++++++++++++

- Add paragraph alignment property (left, right, centered, etc.)
- Add vertical alignment within table cell (top, middle, bottom)
- Add table cell margin properties
- Add table boolean properties: first column (row header), first row (column
  headings), last row (for e.g. totals row), last column (for e.g. row
  totals), horizontal banding, and vertical banding.
- Add support for auto shape adjustment values, e.g. change radius of corner
  rounding on rounded rectangle, position of callout arrow, etc.


0.2.4 (2013-05-16)
++++++++++++++++++

- Add support for auto shapes (e.g. polygons, flowchart symbols, etc.)


0.2.3 (2013-05-05)
++++++++++++++++++

- Add support for table shapes
- Add indentation support to textbox shapes, enabling multi-level bullets on
  bullet slides.


0.2.2 (2013-03-25)
++++++++++++++++++

- Add support for opening and saving a presentation from/to a file-like
  object.
- Refactor XML handling to use lxml objectify


0.2.1 (2013-02-25)
++++++++++++++++++

- Add support for Python 2.6
- Add images from a stream (e.g. StringIO) in addition to a path, allowing
  images retrieved from a database or network resource to be inserted without
  saving first.
- Expand text methods to accept unicode and UTF-8 encoded 8-bit strings.
- Fix potential install bug triggered by importing ``__version__`` from
  package ``__init__.py`` file.


0.2.6 (2013-02-10)
++++++++++++++++++

First non-alpha release with basic capabilities:

- open presentation/template or use built-in default template
- add slide
- set placeholder text (e.g. bullet slides)
- add picture
- add text box
