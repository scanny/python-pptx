.. _table:

Table
=====

One of the shapes available for placing on a PowerPoint slide is the *table*.
As shapes go, it is one of the more complex. In addition to having standard
shape properties such as position and size, it contains what amount to
sub-shapes that each have properties of their own. Prominent among these
sub-element types are row, column, and cell.


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


PowerPoint UI behavior
----------------------

The MS PowerPoint® client exhibits the following behavior related to table
properties:

Upon insertion of a default, empty table
   `<a:tblPr firstRow="1" bandRow="1">` A `tblPr` element is present with a
   `firstRow` and `bandRow` attribute, each set to True.

After setting `firstRow` property off
   `<a:tblPr bandRow="1">` The `firstRow` attribute is removed, not set to
   `0` or `false`.

The `<a:tblPr>` element is always present, even when it contains no
attributes, because it contains an `<a:tableStyleId>` element, even when the
table style is set to none using the UI.


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


MS API Analysis
---------------

MS API method to add a table is::

    Shapes.AddTable(NumRows, NumColumns, Left, Top, Width, Height)

There is a `Shape.HasTable` property which is true when a (`GraphicFrame`)
shape contains a table.

Most interesting `Table` members:

* `Cell(row, col)` method to access individual cells.
* `Columns` collection reference, with `Add` method (`Delete` method is
  on `Column` object)
* `Rows` collection reference
* `FirstCol` and `FirstRow` boolean properties that indicate whether to apply
  special formatting from theme or whatever to first column/row.
* `LastCol`, `LastRow`, and `HorizBanding`, all also boolean with similar
  behaviors.
* `TableStyle` read-only to table style in theme. `Table.ApplyStyle()` method
  is used to set table style.

* `Columns.Add()`
* `Rows.Add()`

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


XML Semantics
-------------

A `tableStyles.xml` part is present in default document, containing the
single (default) style "Medium Style 2 - Accent 1". Colors are specified
indirectly by reference to theme-specified values.


Specimen XML
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. highlight:: xml

Default table produced by PowerPoint::

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


Schema excerpt
--------------

::

  <xsd:element name="tbl" type="CT_Table"/>

  <xsd:complexType name="CT_Table">
    <xsd:sequence>
      <xsd:element name="tblPr"   type="CT_TableProperties" minOccurs="0"/>
      <xsd:element name="tblGrid" type="CT_TableGrid"/>
      <xsd:element name="tr"      type="CT_TableRow"        minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TableProperties">
    <xsd:sequence>
      <xsd:group   ref="EG_FillProperties"   minOccurs="0"/>
      <xsd:group   ref="EG_EffectProperties" minOccurs="0"/>
      <xsd:choice minOccurs="0">
        <xsd:element name="tableStyle"   type="CT_TableStyle"/>
        <xsd:element name="tableStyleId" type="s:ST_Guid"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="rtl"      type="xsd:boolean" default="false"/>
    <xsd:attribute name="firstRow" type="xsd:boolean" default="false"/>
    <xsd:attribute name="firstCol" type="xsd:boolean" default="false"/>
    <xsd:attribute name="lastRow"  type="xsd:boolean" default="false"/>
    <xsd:attribute name="lastCol"  type="xsd:boolean" default="false"/>
    <xsd:attribute name="bandRow"  type="xsd:boolean" default="false"/>
    <xsd:attribute name="bandCol"  type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TableGrid">
    <xsd:sequence>
      <xsd:element name="gridCol" type="CT_TableCol" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TableCol">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="w" type="ST_Coordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TableRow">
    <xsd:sequence>
      <xsd:element name="tc"     type="CT_TableCell"              minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="h" type="ST_Coordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TableCell">
    <xsd:sequence>
      <xsd:element name="txBody" type="CT_TextBody"               minOccurs="0"/>
      <xsd:element name="tcPr"   type="CT_TableCellProperties"    minOccurs="0"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="rowSpan"  type="xsd:int"     default="1"/>
    <xsd:attribute name="gridSpan" type="xsd:int"     default="1"/>
    <xsd:attribute name="hMerge"   type="xsd:boolean" default="false"/>
    <xsd:attribute name="vMerge"   type="xsd:boolean" default="false"/>
    <xsd:attribute name="id"       type="xsd:string"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TextBody">
    <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle" minOccurs="0"/>
      <xsd:element name="p"        type="CT_TextParagraph" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TableCellProperties">
    <xsd:sequence>
      <xsd:element name="lnL"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="lnR"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="lnT"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="lnB"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="lnTlToBr" type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="lnBlToTr" type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="cell3D"   type="CT_Cell3D"                 minOccurs="0"/>
      <xsd:group   ref="EG_FillProperties"                          minOccurs="0"/>
      <xsd:element name="headers"  type="CT_Headers"                minOccurs="0"/>
      <xsd:element name="extLst"   type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="marL"         type="ST_Coordinate32"         default="91440"/>
    <xsd:attribute name="marR"         type="ST_Coordinate32"         default="91440"/>
    <xsd:attribute name="marT"         type="ST_Coordinate32"         default="45720"/>
    <xsd:attribute name="marB"         type="ST_Coordinate32"         default="45720"/>
    <xsd:attribute name="vert"         type="ST_TextVerticalType"     default="horz"/>
    <xsd:attribute name="anchor"       type="ST_TextAnchoringType"    default="t"/>
    <xsd:attribute name="anchorCtr"    type="xsd:boolean"             default="false"/>
    <xsd:attribute name="horzOverflow" type="ST_TextHorzOverflowType" default="clip"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Coordinate">
    <xsd:union memberTypes="ST_CoordinateUnqualified s:ST_UniversalMeasure"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_CoordinateUnqualified">
    <xsd:restriction base="xsd:long">
      <xsd:minInclusive value="-27273042329600"/>
      <xsd:maxInclusive value="27273042316900"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_UniversalMeasure">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)"/>
    </xsd:restriction>
  </xsd:simpleType>


Resources
---------

`Table.FirstCol Property page on MSDN`_

.. _Table.FirstCol Property page on MSDN:
   http://msdn.microsoft.com/en-us/library/office/ff744530.aspx

* ISO-IEC-29500-1, Section 21.1.3 (DrawingML) Tables, pp3331
* ISO-IEC-29500-1, Section 21.1.3.13 tbl (Table), pp3344


.. _Table Members:
   http://msdn.microsoft.com/en-us/library/office/ff745711(v=office.14).aspx

.. _Column Members:
   http://msdn.microsoft.com/en-us/library/office/ff746286(v=office.14).aspx

.. _Cell Members:
   http://msdn.microsoft.com/en-us/library/office/ff744136(v=office.14).aspx

.. _LineFormat Members:
   http://msdn.microsoft.com/en-us/library/office/ff745240(v=office.14).aspx
