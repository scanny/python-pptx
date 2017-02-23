.. _AxisTitle:


Axis Title
==========

A chart axis can have a title. Theoretically, the title text can be drawn
from a cell in the spreadsheet behind the chart; however, there is no
mechanism for specifying this from the PowerPoint 2011 UI (Mac) and there's
no mention of such a procedure I could find on search. So only the "manually
applied" axis title text will be discussed here, other than how to remove any
elements associated with "linked" text, however they may have gotten there.

The title is a rich text container, and can contain arbitrary text with
arbitrary formatting (font, size, color, etc.). There is little but one thing
to distinquish an axis title from an independent text box; its position is
automatically adjusted by the chart to account for resizing and movement.

An axis title is visible whenever present. The only way to "hide" it is to
delete it, along with its contents.

Although it will not be supported, it appears that axis title text can be
specified in the XML as a cell reference in the Excel worksheet. This is
a so-called "linked" title, and in general, any constructive operations on
the axis title will remove any linked title present.


Proposed Scope
--------------

Completed
~~~~~~~~~

* Axis.has_title
* Axis.axis_title
* AxisTitle.has_text_frame
* AxisTitle.text_frame
* AxisTitle.format

Pending
~~~~~~~

* AxisTitle.orientation (this may require text frame rotation)
* XL_ORIENTATION enumeration


Candidate protocol
------------------

Typical usage
~~~~~~~~~~~~~

I expect the most typical use is simply to set the text of the axis title::

    >>> axis = shapes.add_chart(...).chart.value_axis
    >>> axis.axis_title.text_frame.text = 'Foobar'


Axis title access
~~~~~~~~~~~~~~~~~~

``Axis.has_axis_title`` is used to non-destructively test for the presence of
an axis title and may also be used to add an axis title.

An axis on a newly created chart has no axis title::

    >>> axis = shapes.add_chart(...).chart.value_axis
    >>> axis.has_title
    False

Assigning |True| to ``.has_title`` causes an empty axis title element to be
added along with its text frame elements (when not already present)::

    >>> axis.has_title = True
    >>> axis.has_title
    True

``Axis.axis_title`` is used to access the ``AxisTitle`` object for an axis.
It will always return an ``AxisTitle`` object, but it may be destructive in
the sense that it adds an axis title if there is none::

    >>> axis = shapes.add_chart(...).chart.value_axis
    >>> axis.has_title
    False
    >>> axis.axis_title
    <pptx.chart.axis.AxisTitle object at 0x65432fd>
    >>> axis.has_title
    True

Assigning |False| to ``.has_title`` removes the title element from the XML
along with its contents::

    >>> axis.has_title = False
    >>> axis.has_title
    False

Assigning |None| to ``Axis.axis_title`` has the same effect (not sure we'll
actually implement this as a priority)::

    >>> axis.has_title
    True
    >>> axis.axis_title = None
    >>> axis.has_title
    False


AxisTitle.text_frame
~~~~~~~~~~~~~~~~~~~~

According to the schema and the MS API, the ``AxisTitle`` object can contain
either a text frame or an Excel cell reference (``<c:strRef>``). However, the
only operation on the `c:strRef` element the library will support (for now
anyway) is to delete it when adding a text frame.

A newly added axis title will already have a text frame, but for the sake of
completeness, ``AxisTitle.has_text_frame`` will allow the client to test for,
add, and remove an axis title text frame. Assigning |True| to
``.has_text_frame`` causes any Excel reference (``<c:strRef>``) element to be
removed and an empty text frame to be inserted. If a text frame is already
present, no changes are made::

    >>> axis_title.has_text_frame
    False
    >>> axis_title.has_text_frame = True
    >>> axis_title.has_text_frame
    True

The text frame can be accessed using ``AxisTitle.text_frame``. This call
always returns a text frame object, newly created if not already present::

    >>> axis_title.has_text_frame
    False
    >>> axis_title.text_frame
    <pptx.text.text.TextFrame object at 0x65432fe>
    >>> axis_title.has_text_frame
    True


AxisTitle.orientation
~~~~~~~~~~~~~~~~~~~~~

By default, the PowerPoint UI adds an axis title for a vertical axis at 90°
counterclockwise rotation. The MS API provides for rotation to be specified
as an integer number of degrees between -90 and 90. Positive angles are
interpreted as counterclockwise from the horizontal. Orientation can also be
specified as one of the members of the `XlOrientation` enumeration. The
enumeration includes values for horizontal, 90° (upward), -90° (downward),
and (vertically) stacked::

    >>> axis = shapes.add_chart(...).chart.value_axis
    >>> axis_title.orientation
    90
    >>> axis_title.orientation = XL_ORIENTATION.HORIZONTAL
    >>> axis_title.orientation
    0


MS API
--------------

Axis object
~~~~~~~~~~~

Axis.AxisTitle
    Provides access to the AxisTitle object for this axis.

Axis.HasTitle
    Getting indicates presence of axis title. Setting ensures presence or
    absence of axis title.


AxisTitle object
~~~~~~~~~~~~~~~~

AxisTitle.Format
    Provides access to fill and line formatting.

AxisTitle.FormulaLocal
    Returns or sets the cell reference for the axis title text.

AxisTitle.HorizontalAlignment
    Not terrifically useful AFAICT unless title extends to multiple lines.

AxisTitle.IncludeInLayout
    Might not be available via UI; no such option present on PowerPoint 2011
    for Mac.

AxisTitle.Orientation
    An integer value from –90 to 90 degrees or one of the XlOrientation
    constants.

AxisTitle.Text
    Returns or sets the axis title text. Setting removes any existing
    directly-applied formatting, but not title-level formatting.

AxisTitle.VerticalAlignment
    Perhaps not terrifically useful since the textbox is automatically
    positioned and sized, so no difference is visible in the typical cases.


PowerPoint UI Behaviors
-----------------------

* To add an axis title from the PowerPoint UI:

  *Chart Layout (ribbon) > Axis Titles > Vertical Axis Title > Rotated Title*

* The default title "Axis Title" appears when no text has been entered.

* The default orientation of a vertical axis title inserted by the UI is
  rotated 90 degrees counterclockwise. This is initially (before text is
  present) implemented using the `c:txPr` element. That element is removed
  when explicit title text is added.


XlOrientation Enumeration
-------------------------

https://msdn.microsoft.com/en-us/library/office/ff746480.aspx

xlDownward (-4170)
    Text runs downward.

xlHorizontal (-4128)
    Text runs horizontally.

xlUpward (-4171)
    Text runs upward.

xlVertical (-4166)
    Text runs downward and is centered in the cell.


Specimen XML
------------

.. highlight:: xml

Add axis title in UI (but don't set text)::

    <c:valAx>
      <!-- ... -->
      <c:majorGridlines/>

      <c:title>
        <c:layout/>
        <c:overlay val="0"/>
        <c:txPr>
          <a:bodyPr rot="-5400000" vert="horz"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr/>
            </a:pPr>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </c:txPr>
      </c:title>

      <c:numFmt formatCode="General" sourceLinked="1"/>
      <!-- ... -->
    </c:valAx>

Edit text directly in UI. Note that `c:txPr` element is removed when text is
added::

    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr rot="-5400000" vert="horz"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr/>
            </a:pPr>
            <a:r>
              <a:rPr lang="en-US" dirty="0" smtClean="0"/>
              <a:t>Foobar</a:t>
            </a:r>
            <a:endParaRPr lang="en-US" dirty="0"/>
          </a:p>
        </c:rich>
      </c:tx>
      <c:layout/>
      <c:overlay val="0"/>
    </c:title>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:group name="EG_AxShared">
    <xsd:sequence>
      <xsd:element name="axId"           type="CT_UnsignedInt"/>
      <xsd:element name="scaling"        type="CT_Scaling"/>
      <xsd:element name="delete"         type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="axPos"          type="CT_AxPos"/>
      <xsd:element name="majorGridlines" type="CT_ChartLines"        minOccurs="0"/>
      <xsd:element name="minorGridlines" type="CT_ChartLines"        minOccurs="0"/>
      <xsd:element name="title"          type="CT_Title"             minOccurs="0"/>
      <xsd:element name="numFmt"         type="CT_NumFmt"            minOccurs="0"/>
      <xsd:element name="majorTickMark"  type="CT_TickMark"          minOccurs="0"/>
      <xsd:element name="minorTickMark"  type="CT_TickMark"          minOccurs="0"/>
      <xsd:element name="tickLblPos"     type="CT_TickLblPos"        minOccurs="0"/>
      <xsd:element name="spPr"           type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"           type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="crossAx"        type="CT_UnsignedInt"/>
      <xsd:choice minOccurs="0" maxOccurs="1">
        <xsd:element name="crosses"   type="CT_Crosses"/>
        <xsd:element name="crossesAt" type="CT_Double"/>
      </xsd:choice>
    </xsd:sequence>
  </xsd:group>

  <xsd:complexType name="CT_Title">
    <xsd:sequence>
      <xsd:element name="tx"      type="CT_Tx"                minOccurs="0"/>
      <xsd:element name="layout"  type="CT_Layout"            minOccurs="0"/>
      <xsd:element name="overlay" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="spPr"    type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"    type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Tx">
    <xsd:sequence>
      <xsd:choice>
        <xsd:element name="strRef" type="CT_StrRef"/>
        <xsd:element name="rich"   type="a:CT_TextBody"/>
      </xsd:choice>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Layout">
    <xsd:sequence>
      <xsd:element name="manualLayout" type="CT_ManualLayout"  minOccurs="0"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"                type="CT_Transform2D"            minOccurs="0"/>
      <xsd:group   ref ="EG_Geometry"                                          minOccurs="0"/>
      <xsd:group   ref ="EG_FillProperties"                                    minOccurs="0"/>
      <xsd:element name="ln"                  type="CT_LineProperties"         minOccurs="0"/>
      <xsd:group   ref ="EG_EffectProperties"                                  minOccurs="0"/>
      <xsd:element name="scene3d"             type="CT_Scene3D"                minOccurs="0"/>
      <xsd:element name="sp3d"                type="CT_Shape3D"                minOccurs="0"/>
      <xsd:element name="extLst"              type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TextBody">  <!-- text frame -->
    <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle"      minOccurs="0"/>
      <xsd:element name="p"        type="CT_TextParagraph"      maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextBodyProperties">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="prstTxWarp"  type="CT_PresetTextShape"        minOccurs="0"/>
      <xsd:choice minOccurs="0">      <!-- EG_TextAutofit -->
        <xsd:element name="noAutofit"   type="CT_TextNoAutofit"/>
        <xsd:element name="normAutofit" type="CT_TextNormalAutofit"/>
        <xsd:element name="spAutoFit"   type="CT_TextShapeAutofit"/>
      </xsd:choice>
      <xsd:element name="scene3d"     type="CT_Scene3D"                minOccurs="0"/>
      <xsd:choice minOccurs="0">      <!-- EG_Text3D -->
        <xsd:element name="sp3d"        type="CT_Shape3D"/>
        <xsd:element name="flatTx"      type="CT_FlatText"/>
      </xsd:choice>
      <xsd:element name="extLst"      type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="rot"              type="ST_Angle"/>
    <xsd:attribute name="spcFirstLastPara" type="xsd:boolean"/>
    <xsd:attribute name="vertOverflow"     type="ST_TextVertOverflowType"/>
    <xsd:attribute name="horzOverflow"     type="ST_TextHorzOverflowType"/>
    <xsd:attribute name="vert"             type="ST_TextVerticalType"/>
    <xsd:attribute name="wrap"             type="ST_TextWrappingType"/>
    <xsd:attribute name="lIns"             type="ST_Coordinate32"/>
    <xsd:attribute name="tIns"             type="ST_Coordinate32"/>
    <xsd:attribute name="rIns"             type="ST_Coordinate32"/>
    <xsd:attribute name="bIns"             type="ST_Coordinate32"/>
    <xsd:attribute name="numCol"           type="ST_TextColumnCount"/>
    <xsd:attribute name="spcCol"           type="ST_PositiveCoordinate32"/>
    <xsd:attribute name="rtlCol"           type="xsd:boolean"/>
    <xsd:attribute name="fromWordArt"      type="xsd:boolean"/>
    <xsd:attribute name="anchor"           type="ST_TextAnchoringType"/>
    <xsd:attribute name="anchorCtr"        type="xsd:boolean"/>
    <xsd:attribute name="forceAA"          type="xsd:boolean"/>
    <xsd:attribute name="upright"          type="xsd:boolean" default="false"/>
    <xsd:attribute name="compatLnSpc"      type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TextListStyle">
    <xsd:sequence>
      <xsd:element name="defPPr"  type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl1pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl2pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl3pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl4pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl5pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl6pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl7pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl8pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl9pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList"  minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:simpleType name="ST_TextVerticalType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="horz"/>
      <xsd:enumeration value="vert"/>
      <xsd:enumeration value="vert270"/>
      <xsd:enumeration value="wordArtVert"/>
      <xsd:enumeration value="eaVert"/>
      <xsd:enumeration value="mongolianVert"/>
      <xsd:enumeration value="wordArtVertRtl"/>
    </xsd:restriction>
  </xsd:simpleType>
