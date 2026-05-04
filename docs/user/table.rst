
Working with tables
===================

PowerPoint allows text and numbers to be presented in tabular form (aligned
rows and columns) in a reasonably flexible way. A PowerPoint table is not
nearly as functional as an Excel spreadsheet, and is definitely less powerful
than a table in Microsoft Word, but it serves well for most presentation
purposes.


Concepts
--------

There are a few terms worth reviewing as a basis for understanding PowerPoint
tables:

table
  A table is a matrix of cells arranged in aligned rows and columns. This
  orderly arrangement allows a reader to more easily make sense of relatively
  large number of individual items. It is commonly used for displaying
  numbers, but can also be used for blocks of text.

  .. image:: /_static/img/table-01.png
     :scale: 75%

cell
  An individual content "container" within a table. A cell has a text-frame
  in which it holds that content. A PowerPoint table cell can only contain
  text. I cannot hold images, other shapes, or other tables.

  A cell has a background fill, borders, margins, and several other
  formatting settings that can be customized on a cell-by-cell basis.

row
  A side-by-side sequence of cells running across the table, all sharing the
  same top and bottom boundary.

column
  A vertical sequence of cells spanning the height of the table, all sharing
  the same left and right boundary.

table grid, also cell grid
  The underlying cells in a PowerPoint table are strictly regular. In
  a three-by-three table there are nine grid cells, three in each row and
  three in each column. The presence of merged cells can obscure portions of
  the cell grid, but not change the number of cells in the grid. Access to
  a table cell in |pp| is always via that cell's coordinates in the cell
  grid, which may not conform to its visual location (or lack thereof) in the
  table.

merged cell
  A cell can be "merged" with adjacent cells, horizontally, vertically, or
  both, causing the resulting cell to look and behave like a single cell that
  spans the area formerly occupied by those individual cells.

  .. image:: /_static/img/table-02.png
     :scale: 75%

merge-origin cell
  The top-left grid-cell in a merged cell has certain special behaviors. The
  content of that cell is what appears on the slide; content of any "spanned"
  cells is hidden. In |pp| a merge-origin cell can be identified with the
  :attr:`._Cell.is_merge_origin` property. Such a cell can report the size of
  the merged cell with its :attr:`.span_height` and :attr:`.span_width`
  properties, and can be "unmerged" back to its underlying grid cells using
  its :meth:`.split` method.

spanned-cell
  A grid-cell other than the merge-origin cell that is "occupied" by a merged
  cell is called a *spanned cell*. Intuitively, the merge-origin cell "spans"
  the other grid cells within its area. A spanned cell can be identified with
  its :attr:`._Cell.is_spanned` property. A merge-origin cell is not itself
  a spanned cell.


Adding a table
--------------

The following code adds a 3-by-3 table in a new presentation::

    >>> from pptx import Presentation
    >>> from pptx.util import Inches

    >>> # ---create presentation with 1 slide---
    >>> prs = Presentation()
    >>> slide = prs.slides.add_slide(prs.slide_layouts[5])

    >>> # ---add table to slide---
    >>> x, y, cx, cy = Inches(2), Inches(2), Inches(4), Inches(1.5)
    >>> shape = slide.shapes.add_table(3, 3, x, y, cx, cy)

    >>> shape
    <pptx.shapes.graphfrm.GraphicFrame object at 0x1022816d0>
    >>> shape.has_table
    True
    >>> table = shape.table
    >>> table
    <pptx.table.Table object at 0x1096f8d90>

.. image:: /_static/img/table-03.png
   :align: center
   :scale: 60%

A couple things to note:

* :meth:`.SlideShapes.add_table` returns a shape that contains the table, not
  the table itself. In PowerPoint, a table is contained in a graphic-frame
  shape, as is a chart or SmartArt. You can determine whether a shape
  contains a table using its :attr:`~.BaseShape.has_table` property and you
  access the table object using the shape's :attr:`~.GraphicFrame.table`
  property.
* The ``width`` and ``height`` arguments are the **total** width and height
  of the table — not the width of a single column or the height of a single
  row. ``add_table`` distributes ``width`` evenly across the ``cols`` columns
  and ``height`` evenly across the ``rows`` rows, with the last column / row
  absorbing any integer-division remainder so ``sum(col.width) == width`` and
  ``sum(row.height) == height``. See `Table height and row height`_ below for
  the caveat that PowerPoint may later grow individual rows at render time
  to fit their text content.


Adding a table inside a group shape
-----------------------------------

``add_table`` is also available on :class:`.GroupShapes`, so a table can be
authored directly inside a group (issue #627). The group's position and
extents recalculate to include the new table::

    >>> slide = prs.slides.add_slide(prs.slide_layouts[6])
    >>> group = slide.shapes.add_group_shape()
    >>> gf = group.shapes.add_table(
    ...     2, 3, Inches(1), Inches(1), Inches(4), Inches(2)
    ... )
    >>> gf.has_table
    True

The same pattern works on :class:`.LayoutShapes` and :class:`.MasterShapes`
— a layout or master can carry a table as a shared, template-level shape.


Inserting a table into a table placeholder
------------------------------------------

A placeholder allows you to specify the position and size of a shape as part
of the presentation "template", and to place a shape of your choosing into
that placeholder when authoring a presentation based on that template. This
can lead to a better looking presentation, with objects appearing in
a consistent location from slide-to-slide.

Placeholders come in different types, one of which is a *table placeholder*.
A table placeholder behaves like other placeholders except it can only accept
insertion of a table. Other placeholder types accept text bullets or charts.

There is a subtle distinction between a *layout placeholder* and a *slide
placeholder*. A layout placeholder appears in a slide layout, and defines the
position and size of the placeholder "cloned" from it onto each slide created
with that layout. As long as you don't adjust the position or size of the
slide placeholder, it will inherit it's position and size from the layout
placeholder it derives from.

To insert a table into a table placeholder, you need a slide layout that
includes a table placeholder, and you need to create a slide using that
layout. These examples assume that the third slide layout in `template.pptx`
includes a table placeholder::

    >>> prs = Presentation('template.pptx')
    >>> slide = prs.slides.add_slide(prs.slide_layouts[2])

*Accessing the table placeholder.* Generally, the easiest way to access
a placeholder shape is to know its position in the `slide.shapes` collection.
If you always use the same template, it will always show up in the same
position::

    >>> table_placeholder = slide.shapes[1]

*Inserting a table.* A table is inserted into the placeholder by calling its
:meth:`~.TablePlaceholder.insert_table` method and providing the desired
number of rows and columns::

    >>> shape = table_placeholder.insert_table(rows=3, cols=4)

The return value is a |GraphicFrame| shape containing the new table, not the
table object itself. Use the :attr:`~.GraphicFrame.table` property of that
shape to access the table object::

    >>> table = shape.table

The containing shape controls the position and size. Everything else, like
accessing cells and their contents, is done from the table object.


Adding rows and columns
-----------------------

|Table| exposes two one-liner convenience methods for extending a table:
:meth:`.Table.add_row` and :meth:`.Table.add_column`. Each is a thin
wrapper around the collection-level ``add()`` method, so the semantics —
default sizing, cell population, and graphic-frame resizing — match those
described in the following sections::

    >>> from pptx.util import Inches
    >>> # ---append a row, inheriting the last row's height---
    >>> new_row = table.add_row()

    >>> # ---or append a row of explicit height---
    >>> taller_row = table.add_row(height=Inches(1))

    >>> # ---append a column, inheriting the last column's width---
    >>> new_col = table.add_column()

    >>> # ---or append a column of explicit width---
    >>> wider_col = table.add_column(width=Inches(2))


Adding a row to an existing table
---------------------------------

A row can be appended to the bottom of an existing table using the
:meth:`~._RowCollection.add` method on the table's rows collection::

    >>> from pptx.util import Inches
    >>> # ---append a row using the last existing row's height as default---
    >>> new_row = table.rows.add()

    >>> # ---or specify an explicit height---
    >>> taller_row = table.rows.add(height=Inches(1))

The new row is initialized with the same number of cells as the other rows in
the table, each containing a single empty paragraph. The height of the
containing graphic-frame shape is automatically increased by the new row's
height.


Deleting a row from an existing table
-------------------------------------

A row can be removed from a table either by calling :meth:`._Row.delete` on
the row itself, or by passing the row to :meth:`._RowCollection.remove` on
the table's rows collection::

    >>> # ---delete the first row---
    >>> table.rows[0].delete()

    >>> # ---or, equivalently, via the collection---
    >>> table.rows.remove(table.rows[0])

The row's underlying ``a:tr`` element is removed from the table and the height
of the containing graphic-frame shape is reduced accordingly. Subsequent use
of a deleted row object is undefined; most operations on it will raise an
exception.

.. note::

   Deleting a row whose cells participate in a vertical-merge range (as either
   the merge-origin or a spanned cell) may leave the table in an inconsistent
   merge state. If merge integrity matters, split any merged cells that span
   the row first using :meth:`._Cell.split` on the merge-origin cell.


Adding a column to an existing table
------------------------------------

A column can be appended to the right side of an existing table using the
:meth:`~._ColumnCollection.add` method on the table's columns collection::

    >>> from pptx.util import Inches
    >>> # ---append a column using the last existing column's width as default---
    >>> new_column = table.columns.add()

    >>> # ---or specify an explicit width---
    >>> wider_column = table.columns.add(width=Inches(2))

The new column is represented by a new ``a:gridCol`` in the table's
``a:tblGrid`` and a new empty ``a:tc`` cell in every existing row. The width
of the containing graphic-frame shape is automatically increased by the new
column's width.


Deleting a column from an existing table
----------------------------------------

A column can be removed from a table either by calling :meth:`._Column.delete`
on the column itself, or by passing the column to
:meth:`._ColumnCollection.remove` on the table's columns collection::

    >>> # ---delete the first column---
    >>> table.columns[0].delete()

    >>> # ---or, equivalently, via the collection---
    >>> table.columns.remove(table.columns[0])

The column's ``a:gridCol`` element is removed from the table's ``a:tblGrid``
and the ``a:tc`` cell at the same column offset is removed from every row.
The width of the containing graphic-frame shape is reduced accordingly.
Subsequent use of a deleted column object is undefined.

.. note::

   Deleting a column whose cells participate in a horizontal-merge range (as
   either the merge-origin or a spanned cell) may leave the table in an
   inconsistent merge state. If merge integrity matters, split any merged
   cells that span the column first using :meth:`._Cell.split` on the
   merge-origin cell.


Accessing a cell
----------------

All content in a table is in a cell, so getting a reference to one of those
is a good place to start::

    >>> cell = table.cell(0, 0)
    >>> cell.text
    ''
    >>> cell.text = 'Unladen Swallow'

.. image:: /_static/img/table-04.png
   :align: center
   :scale: 60%

The cell is specified by its row, column coordinates as zero-based offsets.
The top-left cell is at row, column (0, 0).

Like an auto-shape, a cell has a text-frame and can contain arbitrary text
divided into paragraphs and runs. Any desired character formatting can be
applied individually to each run.

Often however, cell text is just a simple string. For these cases the
read/write :attr:`._Cell.text` property can be the quickest way to set cell
contents.


Merging cells
-------------

A merged cell is produced by specifying two diagonal cells. The merged cell
will occupy all the grid cells in the rectangular region specified by that
diagonal:

.. image:: /_static/img/table-05.png
   :align: center
   :scale: 60%

::

    >>> cell = table.cell(0, 0)
    >>> other_cell = table.cell(1, 1)
    >>> cell.is_merge_origin
    False
    >>> cell.merge(other_cell)
    >>> cell.is_merge_origin
    True
    >>> cell.is_spanned
    False
    >>> other_cell.is_spanned
    True
    >>> table.cell(0, 1).is_spanned
    True

.. image:: /_static/img/table-06.png
   :align: center
   :scale: 60%

A few things to observe:

* The merged cell appears as a single cell occupying the space formerly
  occupied by the other grid cells in the specified rectangular region.

* The formatting of the merged cell (background color, font etc.) is taken
  from the merge origin cell, the top-left cell of the table in this case.

* Content from the merged cells was migrated to the merge-origin cell. That
  content is no longer present in the spanned grid cells (although you can't
  see those at the moment). The content of each cell appears as a separate
  paragraph in the merged cell; it isn't concatenated into a single
  paragraph. Content is migrated in left-to-right, top-to-bottom order of the
  original cells.

* Calling :attr:`other_cell.merge(cell)` would have the exact same effect. The
  merge origin is always the top-left cell in the specified rectangular
  region. There are four distinct ways to specify a given rectangular region
  (two diagonals, each having two orderings).

* When the merge range spans **every column** of the table (for instance any
  multi-row merge inside a single-column table), ``merge()`` deletes the rows
  beneath the origin row — PowerPoint silently drops such a merge otherwise.
  The origin row absorbs the combined height of the removed rows, so the
  graphic-frame height is unchanged. Symmetrically, a range that spans every
  row collapses into its leftmost column, with the removed columns' widths
  absorbed by the surviving column. Any residual horizontal or vertical merge
  within the surviving row / column is preserved. See issue #636.


Un-merging a cell
-----------------

A merged cell can be restored to its underlying grid cells by calling the
:meth:`~._Cell.split` method on its merge-origin cell. Calling
:meth:`~._Cell.split()` on a cell that is not a merge-origin raises
|ValueError|::

    >>> cell = table.cell(0, 0)
    >>> cell.is_merge_origin
    True
    >>> cell.split()
    >>> cell.is_merge_origin
    False
    >>> table.cell(0, 1).is_spanned
    False

.. image:: /_static/img/table-07.png
   :align: center
   :scale: 60%

Note that the content migration performed as part of the `.merge()` operation
was not reversed.


Table height and row height
---------------------------

The height of a table (and of each row) stored in a `.pptx` file is the
*authored* height — the value written into the XML. PowerPoint treats each row
height as a **minimum**: at render time it will silently grow a row to fit its
text content, including wrapped lines, cell margins, and any font-size
differences. The rendered table on screen may therefore be taller than the sum
of `row.height` values reported by |pp|.

This can be surprising when, for example, positioning a shape immediately
below a table::

    >>> shape = slide.shapes.add_table(rows=3, cols=2, left, top, width, height)
    >>> table = shape.table
    >>> shape.height  # EMU of *authored* height, not rendered height
    914400

When the file is later opened in PowerPoint, the host application performs
its layout pass, grows any rows that need more room, and (on save) writes the
new heights back into the XML. The shape below the table may then overlap the
table's rendered area.

**Why python-pptx can't compute the rendered height.** The final row height
depends on glyph metrics (font face, size, weight, kerning), the line-break
algorithm, cell-margin interactions, and several other inputs that together
comprise PowerPoint's proprietary layout engine. Reproducing that engine
faithfully in pure Python is out of scope for |pp|.

**Workarounds.**

* *Round-trip through PowerPoint or LibreOffice.* The most reliable way to
  obtain accurate row heights is to open the generated `.pptx` in a host
  application (PowerPoint desktop, PowerPoint Online, or LibreOffice Impress)
  and save it back. The application will update every row height to its
  laid-out value. This can be automated on Windows or macOS via COM / AppleScript
  automation of PowerPoint, or on Linux via a headless ``libreoffice --convert-to
  pptx`` invocation.

* *Estimate manually.* If a rough estimate is acceptable, multiply the number
  of text lines you expect in each cell by the font size (converted to EMU via
  :class:`pptx.util.Pt`), add top and bottom cell margins, and take the maximum
  across all cells in the row. This will not match PowerPoint exactly but is
  usually within a few percent for plain text content.

* *Set an explicit minimum height and accept overflow.* For many reporting
  workflows, positioning subsequent shapes with a generous gap below the table
  is simpler than trying to compute the rendered height precisely.

This limitation is tracked as `issue #296`_ and is a known wontfix for the
reasons above.

.. _issue #296: https://github.com/scanny/python-pptx/issues/296


Cell borders
------------

An individual cell can carry explicit borders on each of its four edges and on
each of its two diagonals. Each border is exposed as a |LineFormat| object
reached through one of six properties on the cell:

* ``cell.border_left``      — left edge         (``a:lnL``)
* ``cell.border_right``     — right edge        (``a:lnR``)
* ``cell.border_top``       — top edge          (``a:lnT``)
* ``cell.border_bottom``    — bottom edge       (``a:lnB``)
* ``cell.border_diagonal_down`` — top-left to bottom-right (``a:lnTlToBr``)
* ``cell.border_diagonal_up``   — bottom-left to top-right (``a:lnBlToTr``)

Each |LineFormat| supports the usual color, width, and dash-style settings::

    >>> from pptx.util import Pt
    >>> from pptx.dml.color import RGBColor
    >>> from pptx.enum.dml import MSO_LINE

    >>> cell = table.cell(0, 0)
    >>> cell.border_bottom.color.rgb = RGBColor(0xC0, 0x00, 0x00)  # dark red
    >>> cell.border_bottom.width = Pt(1.5)
    >>> cell.border_bottom.dash_style = MSO_LINE.DASH

Reading a border property reports only values that have been *explicitly*
applied to that edge. Borders inherited from the applied table style
(``Table.style_id``) are not reflected here — PowerPoint resolves those at
render time from its ``tableStyles.xml`` definitions, and python-pptx does not
currently read that part. To force an explicit border to always render, set
color and width on the relevant side. To clear an explicit border, set the
width to ``0`` (which removes the visible line) or assign ``None`` to the
dash style to remove a previously-set preset.

.. note::

   PowerPoint always draws the full rectangular edge of a cell; there is no
   way to draw only part of an edge via ``a:lnL`` / ``a:lnR`` / ``a:lnT`` /
   ``a:lnB``. Diagonal borders (``a:lnTlToBr``, ``a:lnBlToTr``) are drawn
   across the cell's interior, not along an edge.


Applying a table style
----------------------

A PowerPoint table can reference one of the host application's built-in table
styles (e.g. *Medium Style 2 - Accent 1*). The reference is stored as a GUID on
the table XML. |pp| exposes that GUID via :attr:`Table.style_id`, so a style can
be read, changed, or cleared::

    >>> table = shape.table
    >>> # ---default GUID assigned by `add_table()`---
    >>> table.style_id
    '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

    >>> # ---switch to a different built-in style---
    >>> table.style_id = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'

    >>> # ---clear the reference (host falls back to default)---
    >>> table.style_id = None

The GUIDs themselves are not invented by |pp|; PowerPoint (or another host
application) ships a fixed set of built-in styles, each identified by a GUID,
in the presentation's ``tableStyles.xml`` part. To discover the GUIDs
available in a particular template, open the ``.pptx`` in PowerPoint, apply
the style you want to a table, then inspect the ``a:tblPr/a:tableStyleId``
element of that table in the unzipped XML.

.. note::

   |pp| does not currently expose a typed, by-name API for selecting built-in
   table styles (for example, ``table.style = "Medium Style 2 - Accent 1"``);
   that is deferred to a future release. This also means |pp| does not
   validate the assigned GUID against the ``tableStyles.xml`` part — an
   unknown GUID will still be written to the file, but the host application
   will silently fall back to its default style when the file is opened. See
   `issue #27`_ for background and the deferred roadmap.

.. _issue #27: https://github.com/scanny/python-pptx/issues/27


Creating an Excel-styled table
------------------------------

A common request is a table that "looks like an Excel table" — a bold header
row with a solid fill and contrasting font color, alternating light/dark row
banding for readability, and crisp borders around every cell. |pp| has no
single ``make_excel_table()`` call for this, but every ingredient is already
in the public API and they compose cleanly. The recipe below produces a table
matching PowerPoint's built-in *Medium Style 2 - Accent 1* preset, which is
the same blue/white banded look Excel applies by default.

The pieces in play are:

* :attr:`Table.style_id` — assign a built-in PowerPoint table-style GUID so
  that host applications render the table using a familiar preset (see
  `Applying a table style`_).
* :attr:`Table.first_row` / :attr:`Table.horz_banding` — turn on the header
  and alternating-row-band behaviors the style encodes.
* Per-cell :attr:`~._Cell.fill` — paint an explicit header-row background so
  the look does not rely on the host resolving ``tableStyles.xml`` (|pp| does
  not currently read that part).
* Per-run font color and weight via the cell's text-frame — make header text
  white and bold.
* The four ``cell.border_*`` |LineFormat| properties — draw a uniform thin
  border on every cell so the grid is visible regardless of the host's
  style-resolution behavior.

.. rubric:: Copy-pasteable snippet

::

    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    # ---built-in GUID for "Medium Style 2 - Accent 1" (blue)---
    MEDIUM_STYLE_2_ACCENT_1 = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

    HEADER_FILL = RGBColor(0x4F, 0x81, 0xBD)   # Accent 1 blue
    HEADER_FONT = RGBColor(0xFF, 0xFF, 0xFF)   # white
    BORDER      = RGBColor(0x95, 0xB3, 0xD7)   # light accent-1 border

    def style_as_excel_table(table, headers, rows):
        """Populate *table* with *headers* and *rows* and apply Excel-like styling.

        *table* must already have ``len(rows) + 1`` rows and ``len(headers)``
        columns. *headers* is a sequence of strings. *rows* is a sequence of
        equal-length sequences of stringifiable values.
        """
        # ---ask the host to render the familiar Medium Style 2 - Accent 1 preset---
        table.style_id = MEDIUM_STYLE_2_ACCENT_1
        table.first_row = True        # bold, filled header row
        table.horz_banding = True     # alternating row bands

        # ---populate header row and paint it explicitly---
        for col_idx, label in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = label
            cell.fill.solid()
            cell.fill.fore_color.rgb = HEADER_FILL
            run = cell.text_frame.paragraphs[0].runs[0]
            run.font.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = HEADER_FONT

        # ---populate body rows---
        for row_idx, row in enumerate(rows, start=1):
            for col_idx, value in enumerate(row):
                cell = table.cell(row_idx, col_idx)
                cell.text = str(value)
                run = cell.text_frame.paragraphs[0].runs[0]
                run.font.size = Pt(11)

        # ---draw a thin, uniform border on every cell so the grid is visible
        # regardless of how the host resolves tableStyles.xml---
        for cell in table.iter_cells():
            for border in (
                cell.border_left,
                cell.border_right,
                cell.border_top,
                cell.border_bottom,
            ):
                border.color.rgb = BORDER
                border.width = Pt(0.75)

    # ---build the presentation---
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    headers = ["Region", "Q1", "Q2", "Q3", "Q4"]
    data = [
        ("North",  12500, 13800, 14200, 15100),
        ("South",   9800, 10400, 11100, 11900),
        ("East",   15200, 15900, 16400, 17200),
        ("West",   11100, 11700, 12000, 12800),
    ]

    shape = slide.shapes.add_table(
        rows=len(data) + 1,
        cols=len(headers),
        left=Inches(1), top=Inches(1),
        width=Inches(8), height=Inches(2.5),
    )
    style_as_excel_table(shape.table, headers, data)

    prs.save("excel_styled_table.pptx")

A few notes on the recipe:

* The GUID is just a string — any of PowerPoint's built-in table-style GUIDs
  will do. Swap in ``"{2D5ABB26-0587-4C30-8999-92F81FD0307C}"`` (Medium Style
  2 - Accent 2) or any other to recolor the preset. See
  `Applying a table style`_ for how to discover GUIDs in a particular
  template.
* The explicit ``fill`` and ``border_*`` settings are defensive: because |pp|
  does not currently read ``tableStyles.xml``, anything the host is "supposed"
  to paint from the referenced style may not render if the file is opened in
  a reader that lacks that style part. Painting header fill and cell borders
  explicitly guarantees a consistent look everywhere the file is opened.
* ``table.horz_banding = True`` turns on the banding *bit* but the band colors
  themselves come from the referenced table style. If you need exact,
  host-independent band colors, set ``cell.fill`` explicitly on every even-row
  cell as well (mirroring the header-row loop above).
* Adding rows after the table exists — for example, when populating from
  a query of unknown length — works the same way: call
  :meth:`._RowCollection.add` to append, then run the same per-cell styling
  loop on the new cells. See `Adding a row to an existing table`_.


A few snippets that might be handy
----------------------------------

Use Case: Interrogate table for merged cells::

    def iter_merge_origins(table):
        """Generate each merge-origin cell in *table*.

        Cell objects are ordered by their position in the table,
        left-to-right, top-to-bottom.
        """
        return (cell for cell in table.iter_cells() if cell.is_merge_origin)

    def merged_cell_report(cell):
        """Return str summarizing position and size of merged *cell*."""
        return (
            'merged cell at row %d, col %d, %d cells high and %d cells wide'
            % (cell.row_idx, cell.col_idx, cell.span_height, cell.span_width)
        )

    # ---Print a summary line for each merged cell in *table*.---
    for merge_origin_cell in iter_merge_origins(table):
        print(merged_cell_report(merge_origin_cell))

prints a report like::

    merged cell at row 0, col 0, 2 cells high and 2 cells wide
    merged cell at row 3, col 2, 1 cells high and 2 cells wide
    merged cell at row 4, col 0, 2 cells high and 1 cells wide

Use Case: Access only cells that display text (are not spanned)::

    def iter_visible_cells(table):
        return (cell for cell in table.iter_cells() if not cell.is_spanned)

Use Case: Determine whether table contains merged cells::

    def has_merged_cells(table):
        for cell in table.iter_cells():
            if cell.is_merge_origin:
                return True
        return False
