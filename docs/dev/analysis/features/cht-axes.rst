
Chart Axes
==========

PowerPoint chart axes come in four varieties: category axis, value axis, date
axis, and series axis. A series axis only appears on a 3D chart and is also
known as its depth axis.

A chart may have two category axes and/or two value axes. The second axis, if
there is one, is known as the *secondary category axis* or *secondary value
axis*.

A category axis may appear as either the horizontal or vertical axis,
depending upon the chart type. Likewise for a value axis.


PowerPoint behavior
-------------------

Tick label position
~~~~~~~~~~~~~~~~~~~

Proposed python-pptx protocol::

    >>> axis.tick_label_position
    XL_TICK_LABEL_POSITION.NEXT_TO_AXIS
    >>> axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    >>> axis.tick_label_position
    XL_TICK_LABEL_POSITION.LOW

MS API protocol::

    >>> axis.TickLabelPosition
    xlTickLabelPositionNextToAxis
    >>> axis.TickLabelPosition = xlTickLabelPositionLow

    c:catAx/c:tickLblPos{val=nextTo}

Option "none" causes tick labels to be hidden.

Default when no ``<c:tickLblPos>`` element is present is nextTo. Same if
element is present with no ``val`` attribute.


TickLabels.number_format
~~~~~~~~~~~~~~~~~~~~~~~~

Proposed python-pptx protocol::

    >>> tick_labels = axis.tick_labels
    >>> tick_labels.number_format
    'General'
    >>> tick_labels.number_format_is_linked
    True
    >>> tick_labels.number_format = '#,##0.00'
    >>> tick_labels.number_format_is_linked = False

Tick-mark label text for the category axis comes from the name of the
associated category in the chart. The default tick-mark label text for the
category axis is the number that indicates the position of the category
relative to the left end of this axis. To change the number of unlabeled tick
marks between tick-mark labels, you must change the TickLabelSpacing property
for the category axis.

Tick-mark label text for the value axis is calculated based on the MajorUnit,
MinimumScale, and MaximumScale properties of the value axis. To change the
tick-mark label text for the value axis, you must change the values of these
properties.

MS API protocol::

    >>> tick_labels = axis.TickLabels
    >>> tick_labels.NumberFormat
    'General'
    >>> tick_labels.NumberFormatLinked
    True
    >>> tick_labels.NumberFormatLinked = False
    >>> tick_labels.NumberFormat = "#,##0.00"

    # output of 1234.5678 is '1,234.57'

When ``sourceLinked`` attribute is True, UI shows "General" number format
category regardless of contents of ``formatCode`` attribute.

The ``sourceLinked`` attribute defaults to True when the ``<c:numFmt>``
element is present but that attribute is omitted.

When the ``<c:numFmt>`` element is not present, the behavior is as though the
element ``<c:numFmt formatCode="General" sourceLinked="0"/>`` was present.

The default PowerPoint chart contains this numFmt element::

    ``<c:numFmt formatCode="General" sourceLinked="1"/>``.


_BaseAxis.visible property
~~~~~~~~~~~~~~~~~~~~~~~~~~

``<c:delete val="0"/>`` element

* when delete element is absent, the default value is True
* when ``val`` attribute is absent, the default value is True


XML specimens
-------------

.. highlight:: xml

Example axis XML for a single-series line plot::

  <c:catAx>
    <c:axId val="-2097691448"/>
    <c:scaling>
      <c:orientation val="minMax"/>
    </c:scaling>
    <c:delete val="0"/>
    <c:axPos val="b"/>
    <c:numFmt formatCode="General" sourceLinked="0"/>
    <c:majorTickMark val="out"/>
    <c:minorTickMark val="none"/>
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
    <c:auto val="1"/>
    <c:lblAlgn val="ctr"/>
    <c:lblOffset val="100"/>
    <c:noMultiLvlLbl val="0"/>
  </c:catAx>


Related Schema Definitions
--------------------------

::

  <xsd:complexType name="CT_PlotArea">
    <xsd:sequence>
      <!-- 17 others -->
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="valAx"  type="CT_ValAx"/>
        <xsd:element name="catAx"  type="CT_CatAx"/>
        <xsd:element name="dateAx" type="CT_DateAx"/>
        <xsd:element name="serAx"  type="CT_SerAx"/>
      </xsd:choice>
      <xsd:element name="dTable" type="CT_DTable"            minOccurs="0"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_CatAx">  <!-- denormalized -->
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
      <xsd:choice                                                    minOccurs="0">
        <xsd:element name="crosses"      type="CT_Crosses"/>
        <xsd:element name="crossesAt"    type="CT_Double"/>
      </xsd:choice>
      <xsd:element name="auto"           type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="lblAlgn"        type="CT_LblAlgn"           minOccurs="0"/>
      <xsd:element name="lblOffset"      type="CT_LblOffset"         minOccurs="0"/>
      <xsd:element name="tickLblSkip"    type="CT_Skip"              minOccurs="0"/>
      <xsd:element name="tickMarkSkip"   type="CT_Skip"              minOccurs="0"/>
      <xsd:element name="noMultiLvlLbl"  type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ValAx">  <!-- denormalized -->
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
      <xsd:choice                                                    minOccurs="0">
        <xsd:element name="crosses"   type="CT_Crosses"/>
        <xsd:element name="crossesAt" type="CT_Double"/>
      </xsd:choice>
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

  <xsd:complexType name="CT_Scaling">
    <xsd:sequence>
      <xsd:element name="logBase"     type="CT_LogBase"       minOccurs="0"/>
      <xsd:element name="orientation" type="CT_Orientation"   minOccurs="0"/>
      <xsd:element name="max"         type="CT_Double"        minOccurs="0"/>
      <xsd:element name="min"         type="CT_Double"        minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_NumFmt">
    <xsd:attribute name="formatCode"   type="xsd:string"  use="required"/>
    <xsd:attribute name="sourceLinked" type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TickLblPos">
    <xsd:attribute name="val" type="ST_TickLblPos" default="nextTo"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TickMark">
    <xsd:attribute name="val" type="ST_TickMark" default="cross"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Boolean">
    <xsd:attribute name="val" type="xsd:boolean" default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Double">
    <xsd:attribute name="val" type="xsd:double" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_TickLblPos">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="high"/>
      <xsd:enumeration value="low"/>
      <xsd:enumeration value="nextTo"/>
      <xsd:enumeration value="none"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TickMark">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="cross"/>
      <xsd:enumeration value="in"/>
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="out"/>
    </xsd:restriction>
  </xsd:simpleType>
