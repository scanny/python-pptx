
Chart - Tick Labels
===================

The vertical and horizontal divisions of a chart axis may be labeled with
*tick labels*, text that describes the division, most commonly its category
name or value. A tick labels object is not a collection. There is no object
that represents in individual tick label.

Tick label text for a category axis comes from the name of each category. The
default tick label text for a category axis is the number that indicates the
position of the category relative to the low end of this axis. To change the
number of unlabeled tick marks between tick-mark labels, you must change the
TickLabelSpacing property for the category axis.

Tick label text for the value axis is calculated based on the `major_unit`,
`minimum_scale`, and `maximum_scale` properties of the value axis. To change
the tick label text for the value axis, you must change the values of these
properties.


Candidate protocol
------------------

::

    >>> tick_labels = value_axis.tick_labels

    >>> tick_labels.font
    <pptx.text.Font object at 0xdeadbeef1>

    >>> tick_labels.number_format
    'General'
    >>> tick_labels.number_format = '0"%"'
    >>> tick_labels.number_format
    '0"%"'

    >>> tick_labels.number_format_is_linked
    True
    >>> tick_labels.number_format_is_linked = False
    >>> tick_labels.number_format_is_linked
    False

    # offset property is only available on category axis
    >>> tick_labels = category_axis.tick_labels
    >>> tick_labels.offset
    100
    >>> tick_labels.offset = 250
    >>> tick_labels.offset
    250


Feature Summary
---------------

* **TickLabels.font** -- Read/only Font object for tick labels.
* **TickLabels.number_format** -- Read/write string.
* **TickLabels.number_format_is_linked** -- Read/write boolean.
* **TickLabels.offset** -- Read/write int between 0 and 1000, inclusive.


Microsoft API
-------------

Font
    Returns the font of the specified object. Read-only ChartFont.

NumberFormat
    Returns or sets the format code for the object. Read/write String.

NumberFormatLinked
    True if the number format is linked to the cells (so that the number
    format changes in the labels when it changes in the cells). Read/write
    Boolean.

Offset
    Returns or sets the distance between the levels of labels, and the
    distance between the first level and the axis line. Read/write Long.


XML specimens
-------------

.. highlight:: xml

Example category axis XML showing tick label-related elements::

  <c:catAx>
    <c:axId val="-2097691448"/>
    <c:scaling>
      <c:orientation val="minMax"/>
    </c:scaling>
    <c:axPos val="b"/>
    <c:numFmt formatCode="General" sourceLinked="0"/>
    <c:tickLblPos val="nextTo"/>
    <c:txPr>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
        <a:pPr>
          <a:defRPr sz="1000"/>
        </a:pPr>
        <a:endParaRPr lang="en-US"/>
      </a:p>
    </c:txPr>
    <c:crossAx val="-2097683336"/>
    <c:crosses val="autoZero"/>
    <c:lblAlgn val="ctr"/>
    <c:lblOffset val="100"/>
    <c:noMultiLvlLbl val="0"/>
  </c:catAx>


Related Schema Definitions
--------------------------

::

  <xsd:complexType name="CT_CatAx">
    <xsd:sequence>
      <xsd:group    ref="EG_AxShared"/>
      <xsd:element name="auto"           type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="lblAlgn"        type="CT_LblAlgn"           minOccurs="0"/>
      <xsd:element name="lblOffset"      type="CT_LblOffset"         minOccurs="0"/>
      <xsd:element name="tickLblSkip"    type="CT_Skip"              minOccurs="0"/>
      <xsd:element name="tickMarkSkip"   type="CT_Skip"              minOccurs="0"/>
      <xsd:element name="noMultiLvlLbl"  type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ValAx">
    <xsd:sequence>
      <xsd:group    ref="EG_AxShared"/>
      <xsd:element name="crossBetween"   type="CT_CrossBetween"      minOccurs="0"/>
      <xsd:element name="majorUnit"      type="CT_AxisUnit"          minOccurs="0"/>
      <xsd:element name="minorUnit"      type="CT_AxisUnit"          minOccurs="0"/>
      <xsd:element name="dispUnits"      type="CT_DispUnits"         minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_DateAx">
    <xsd:sequence>
      <xsd:group    ref="EG_AxShared"/>
      <xsd:element name="auto"          type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="lblOffset"     type="CT_LblOffset"     minOccurs="0"/>
      <xsd:element name="baseTimeUnit"  type="CT_TimeUnit"      minOccurs="0"/>
      <xsd:element name="majorUnit"     type="CT_AxisUnit"      minOccurs="0"/>
      <xsd:element name="majorTimeUnit" type="CT_TimeUnit"      minOccurs="0"/>
      <xsd:element name="minorUnit"     type="CT_AxisUnit"      minOccurs="0"/>
      <xsd:element name="minorTimeUnit" type="CT_TimeUnit"      minOccurs="0"/>
      <xsd:element name="extLst"        type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SerAx">
    <xsd:sequence>
      <xsd:group    ref="EG_AxShared"/>
      <xsd:element name="tickLblSkip"  type="CT_Skip"          minOccurs="0"/>
      <xsd:element name="tickMarkSkip" type="CT_Skip"          minOccurs="0"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

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

  <xsd:complexType name="CT_LblOffset">
    <xsd:attribute name="val" type="ST_LblOffset" default="100%"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_LblOffset">
    <xsd:union memberTypes="ST_LblOffsetPercent ST_LblOffsetUShort"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_LblOffsetPercent">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="0*(([0-9])|([1-9][0-9])|([1-9][0-9][0-9])|1000)%"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_LblOffsetUShort">
    <xsd:restriction base="xsd:unsignedShort">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="1000"/>
    </xsd:restriction>
  </xsd:simpleType>
