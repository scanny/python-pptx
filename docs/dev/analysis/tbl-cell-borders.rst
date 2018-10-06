Table Cell Borders
------------------

In PowerPoint, a table cell can have borders on all 4 sides, plus two diagonal lines (bottom left to top right and
top left to bottom right) for a total of 6 possible borders, each specified by a child element of the table cell
properties element.

Each border is further customized by a set of attributes and child elements common to all borders (``LineFormat``).

In order to provide ``LineFormat`` functionality, each border ``_ln`` object is wrapped in a ``_CellBorderLine``,
which implements the ``get_or_add_ln`` method, and can be then wrapped by ``LineFormat``.

Protocol
--------

**Starting with a default cell** A new cell has no borders or style.
New borders can be created by accessing the borders property::

    >>> table = shapes.add_table(rows, columns, left, top, width, height).table
    >>> cell = table.rows[0].cells[0]
    >>> cell
    <pptx.shapes.table._Cell object at 0x...>
    >>> cell.borders
    <pptx.shapes.table._CellBorders object at 0x...>

**Create a border** A cell border is created once it's accessed for the first time.
It won't be visible, however, until applying line width and color for its ``LineFormat``.
This can be done with any one of the 6 cell borders::

    >>> cell.borders.top
    <pptx.dml.line.LineFormat at 0x...>
    >>> cell.borders.top.width
    0
    >>> cell.borders.top.width = Pt(1)
    >>> cell.borders.top.width
    12700
    >>> cell.borders.top.color.rgb = RGBColor(0, 0, 0)

**Apply different line styles** Cell border work with ``LineFormat``, so any other line style can also be applied::

    >>> cell.borders.top.dash_style = MSO_LINE_DASH_STYLE.DASH_DOT_DOT


Scope
-----

* `_Cell.borders` -> `_CellBorders` - An immutable collection of 6 `LineFormat` objects
* `_CellBorderLine` is used as a child object for `LineFormat`
* `_CellBorders.top` -> `LineFormat`
* ... other borders (bottom, left, right, top-left to bottom-right, bottom-left to top-right)
* No option to clear border (other than to give its `LineFormat` a width of 0)

MS API
------

* | MS API supplies a ``Borders`` object to enable easy access to all cell borders
  | https://docs.microsoft.com/en-us/office/vba/api/powerpoint.cell.borders

* | The borders collection object holds 6 borders (top, bottom, left, right, diagonal down and diagonal up),
  | and cannot be modified
  | https://docs.microsoft.com/en-us/office/vba/api/powerpoint.borders

* | Table cell borders object is actually a collection of ``LineFormat`` objects, one for each border
  | https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/lineformat-object-powerpoint

Specimen XML
------------

Default table cell::

    <a:tc>
      <a:txBody .../>
      <a:tcPr/>
    </a:tc>

Table cell with top border::

    <a:tc>
      <a:txBody .../>
      <a:tcPr>
        <a:lnT w="12700"/>
      </a:tcPr>
    </a:tc>

Table cell with 4 borders::

    <a:tc>
      <a:txBody .../>
      <a:tcPr>
        <a:lnL w="12700"/>
        <a:lnR w="12700"/>
        <a:lnT w="12700"/>
        <a:lnB w="12700"/>
      </a:tcPr>
    </a:tc>


Related Schema Definitions
--------------------------

::

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

  <xsd:complexType name="CT_LineProperties">
    <xsd:sequence>
      <xsd:group   ref="EG_LineFillProperties"                     minOccurs="0"/>
      <xsd:group   ref="EG_LineDashProperties"                     minOccurs="0"/>
      <xsd:group   ref="EG_LineJoinProperties"                     minOccurs="0"/>
      <xsd:element name="headEnd" type="CT_LineEndProperties"      minOccurs="0"/>
      <xsd:element name="tailEnd" type="CT_LineEndProperties"      minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="w"    type="ST_LineWidth"/>
    <xsd:attribute name="cap"  type="ST_LineCap"/>
    <xsd:attribute name="cmpd" type="ST_CompoundLine"/>
    <xsd:attribute name="algn" type="ST_PenAlignment"/>
  </xsd:complexType>
