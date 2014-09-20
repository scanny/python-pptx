
Table
=====


Overview
--------

One of the shapes available for placing on a PowerPoint slide is the *table*.
As shapes go, it is one of the more complex. In addition to having standard
shape properties such as position and size, it contains what amount to
sub-shapes that each have properties of their own. Prominent among these
sub-element types are row, column, and cell.


Open questions
--------------

* not sure of the semantics of the ``<a:tableStyleId>`` element. Assuming the
  GUID it contains somehow maps to a style in the tableStyles.xml, or perhaps
  some contant values defined elsewhere.

* What would be the Pythonic way to delete a row or column from a table? The MS
  API has a ``delete()`` method on the row or column itself. I'm inclined to
  believe it would be more Pythonic to use the ``del`` keyword on the indexed
  list, e.g. ``del rows[2]``.


Table Properties
----------------

``first_row``
   read/write boolean property which, when true, indicates the first row should
   be formatted differently, as for a heading row at the top of the table.

``first_col``
   read/write boolean property which, when true, indicates the first column
   should be formatted differently, as for a side-heading column at the far
   left of the table.

``last_row``
   read/write boolean property which, when true, indicates the last row should
   be formatted differently, as for a totals row at the bottom of the table.

``last_col``
   read/write boolean property which, when true, indicates the last column
   should be formatted differently, as for a side totals column at the far
   right of the table.

``horz_banding``
   read/write boolean property indicating whether alternate color "banding"
   should be applied to the body rows.

``vert_banding``
   read/write boolean property indicating whether alternate color "banding"
   should be applied to the table columns.


Attribute Names
~~~~~~~~~~~~~~~

+---------------+-----------+-------------+
| property name | attribute | optionality |
+===============+===========+=============+
| first_row     | firstRow  | optional    |
+---------------+-----------+-------------+
| first_col     | firstCol  | optional    |
+---------------+-----------+-------------+
| last_row      | lastRow   | optional    |
+---------------+-----------+-------------+
| last_col      | lastCol   | optional    |
+---------------+-----------+-------------+
| horz_banding  | bandRow   | optional    |
+---------------+-----------+-------------+
| vert_banding  | bandCol   | optional    |
+---------------+-----------+-------------+


Characteristics
~~~~~~~~~~~~~~~

+------------+---------------+
| value type | boolean, None |
+------------+---------------+
| mode       | read/write    |
+------------+---------------+


XML location
~~~~~~~~~~~~

.. highlight:: xml

::

   <a:tbl>
     <a:tblPr firstRow="1" lastCol="1" bandRow="1">

``<a:tblPr>`` is an optional element. If it appears, it is the first child of
``<a:tbl>``


Observed behavior
~~~~~~~~~~~~~~~~~

The MS PowerPoint® client exhibits the following behavior related to table
properties:

upon insertion of a default, empty table
   ``<a:tblPr firstRow="1" bandRow="1">`` A tblPr element is present with a
   ``firstRow`` and ``bandRow`` attribute, each set to True.

after setting ``firstRow`` property off
   ``<a:tblPr bandRow="1">`` The ``firstRow`` attribute is removed, not set
   to ``0`` or ``false``.

The ``<a:tblPr>`` element is always present, even when it contains no
attributes, because it contains an ``<a:tableStyleId>`` element, even when
the table style is set to none using the UI.


References
~~~~~~~~~~

`Table.FirstCol Property page on MSDN`_

.. _Table.FirstCol Property page on MSDN:
   http://msdn.microsoft.com/en-us/library/office/ff744530.aspx


API requirements
----------------

|Table| class
~~~~~~~~~~~~~

Properties and methods required for a |Table| shape.

* ``apply_style(style_id)`` -- change the style of the table. Not sure what the
  domain of ``style_id`` is.

* ``cell(row, col)`` -- method to access an individual |_Cell| object.

* ``columns`` -- collection of |_Column| objects, each corresponding to
  a column in the table, in left-to-right order.

* ``first_col`` -- read/write boolean property which, when true, indicates the
  first column should be formatted differently, as for a side-heading column at
  the far left of the table.

* ``first_row`` -- read/write boolean property which, when true, indicates the
  first row should be formatted differently, as for a heading row at the top of
  the table.

* ``horz_banding`` -- read/write boolean property indicating whether alternate
  color "banding" should be applied to the body rows.

* ``last_col`` -- read/write boolean property which, when true, indicates the
  last column should be formatted differently, as for a side totals column at
  the far right of the table.

* ``last_row`` -- read/write boolean property which, when true, indicates the
  last row should be formatted differently, as for a totals row at the bottom
  of the table.

* ``rows`` -- collection of |_Row| objects, each corresponding to a row in the
  table, in top-to-bottom order.

* ``vert_banding`` -- read/write boolean property indicating whether alternate
  color "banding" should be applied to the table columns.


|_Cell| class
~~~~~~~~~~~~~

* ``text_frame`` -- container for text in the cell.
* borders, something like LineProperties on each side
* inset (margins)
* anchor and anchor_center
* horzOverflow, not sure what this is exactly, maybe wrap or auto-resize to
  fit.


|_Column| class
~~~~~~~~~~~~~~~

Provide the properties and methods appropriate to a table column.

* ``width`` -- read/write integer width of the column in English Metric Units
* perhaps ``delete()`` method


|_ColumnCollection| class
~~~~~~~~~~~~~~~~~~~~~~~~~

* ``add(before)`` -- add a new column to the left of the column having index
  *before*, returning a reference to the new column. *before* defaults to
  ``-1``, which adds the column as the last column in the table.


|_Row| class
~~~~~~~~~~~~

* ``height`` -- read/write integer height of the row in English Metric Units
  (EMU).


|_RowCollection| class
~~~~~~~~~~~~~~~~~~~~~~

* ``add(before)`` -- add a new row before the row having index *before*,
  returning a reference to the new row. *before* defaults to ``-1``, which adds
  the row as the last row in the table.


Behavior
--------

Table width and column widths
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

A table is created by specifying a row and column count, a position, and an
overall size. Initial column widths are set by dividing the overall width by
the number of columns, resolving any rounding errors in the last column.
Conversely, when a column's width is specified, the table width is adjusted to
the sum of the widths of all columns. Initial row heights are set similarly and
overall table height adjusted to the sum of row heights when a row's height is
specified.


Discovery protocol
------------------

* (/) Review MS API documentation
* (/) Inspect minimal XML produced by PowerPoint® client
* (.) Review and document relevant schema elements


MS API Analysis
---------------

MS API method to add a table is::

    Shapes.AddTable(NumRows, NumColumns, Left, Top, Width, Height)

There is a HasTable property on Shape to indicate the shape "has" a table.
Seems like "is" a table would be more apt, but I'm still looking :)

From the `Table Members`_ page on MSDN.

Most interesting ``Table`` members:

* ``Cell(row, col)`` method to access individual cells.
* ``Columns`` collection reference, with ``Add`` method (``Delete`` method is
  on ``Column`` object)
* ``Rows`` collection reference
* FirstCol and FirstRow boolean properties that indicate whether to apply
  special formatting from theme or whatever to first column/row.
* LastCol, LastRow, and HorizBanding, all also boolean with similar behaviors
* TableStyle read-only to table style in theme. Table.ApplyStyle() method is
  used to set table style.

Columns collection and Rows collection both have an Add() method

`Column Members`_ page on MSDN.

* Delete()
* Width property

`Cell Members`_ page on MSDN.

* Merge() and Split() methods
* Borders reference to Borders collection of LineFormat objects
* Shape reference to shape object that cell is or has.

`LineFormat Members`_ page on MSDN.

* ForeColor
* Weight


XML produced by PowerPoint® application
---------------------------------------

Inspection Notes
~~~~~~~~~~~~~~~~

A ``tableStyles.xml`` part is fleshed out substantially; looks like it's
populated from built-in defaults "Medium Style 2 - Accent 1". It appears to
specify colors indirectly by reference to theme-specified values.


XML produced by PowerPoint® client
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. highlight:: xml

::

    <p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="2" name="Table 1"/>
        <p:cNvGraphicFramePr>
          <a:graphicFrameLocks noGrp="1"/>
        </p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="1524000" y="1397000"/>
        <a:ext cx="6096000" cy="741680"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr firstRow="1" bandRow="1">
              <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
            </a:tblPr>
            <a:tblGrid>
              <a:gridCol w="3048000"/>
              <a:gridCol w="3048000"/>
            </a:tblGrid>
            <a:tr h="370840">
              <a:tc>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:endParaRPr lang="en-US"/>
                  </a:p>
                </a:txBody>
                <a:tcPr/>
              </a:tc>
              <a:tc>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:endParaRPr lang="en-US"/>
                  </a:p>
                </a:txBody>
                <a:tcPr/>
              </a:tc>
            </a:tr>
            <a:tr h="370840">
              <a:tc>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:endParaRPr lang="en-US"/>
                  </a:p>
                </a:txBody>
                <a:tcPr/>
              </a:tc>
              <a:tc>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:endParaRPr lang="en-US"/>
                  </a:p>
                </a:txBody>
                <a:tcPr/>
              </a:tc>
            </a:tr>
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>



.. _Table Members:
   http://msdn.microsoft.com/en-us/library/office/ff745711(v=office.14).aspx

.. _Column Members:
   http://msdn.microsoft.com/en-us/library/office/ff746286(v=office.14).aspx

.. _Cell Members:
   http://msdn.microsoft.com/en-us/library/office/ff744136(v=office.14).aspx

.. _LineFormat Members:
   http://msdn.microsoft.com/en-us/library/office/ff745240(v=office.14).aspx

