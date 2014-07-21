
Bar Chart
=========

A bar chart is one of the fundamental plot types, used for column and bar
charts, clustered, stacked, and stacked 100%.


Gap Width
---------

A gap appears between the bar or clustered bars for each category on a bar
chart. The default width for this gap is 150% of the bar width. It can be set
between 0 and 500% of the bar width. In the MS API this is set using the
property `ChartGroup.GapWidth`.

Proposed protocol::

    >>> assert isinstance(bar_plot, BarPlot)
    >>> bar_plot.gap_width
    150
    >>> bar_plot.gap_width = 300
    >>> bar_plot.gap_width
    300
    >>> bar_plot.gap_width = 700
    ValueError: gap width must be in range 0-500 (percent)


XML specimens
-------------

.. highlight:: xml

Minimal working XML for a single-series column plot. Note this does not
reference values in a spreadsheet.::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:chart>
      <c:plotArea>
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:cat>
              <c:strLit>
                <c:ptCount val="1"/>
                <c:pt idx="0">
                  <c:v>Foo</c:v>
                </c:pt>
              </c:strLit>
            </c:cat>
            <c:val>
              <c:numLit>
                <c:ptCount val="1"/>
                <c:pt idx="0">
                  <c:v>4.3</c:v>
                </c:pt>
              </c:numLit>
            </c:val>
          </c:ser>
          <c:dLbls>
            <c:showLegendKey val="0"/>
            <c:showVal val="0"/>
            <c:showCatName val="0"/>
            <c:showSerName val="0"/>
            <c:showPercent val="0"/>
            <c:showBubbleSize val="0"/>
          </c:dLbls>
          <c:gapWidth val="300"/>
          <c:axId val="-2068027336"/>
          <c:axId val="-2113994440"/>
        </c:barChart>
        <c:catAx>
          <c:axId val="-2068027336"/>
          <c:scaling/>
          <c:delete val="0"/>
          <c:axPos val="b"/>
          <c:crossAx val="-2113994440"/>
        </c:catAx>
        <c:valAx>
          <c:axId val="-2113994440"/>
          <c:scaling/>
          <c:delete val="0"/>
          <c:axPos val="l"/>
          <c:crossAx val="-2068027336"/>
        </c:valAx>
      </c:plotArea>
    </c:chart>
  </c:chartSpace>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_BarChart">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="barDir"     type="CT_BarDir"/>
      <xsd:element name="grouping"   type="CT_BarGrouping"   minOccurs="0"/>
      <xsd:element name="varyColors" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="ser"        type="CT_BarSer"        minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"      type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="gapWidth"   type="CT_GapAmount"     minOccurs="0"/>
      <xsd:element name="overlap"    type="CT_Overlap"       minOccurs="0"/>
      <xsd:element name="serLines"   type="CT_ChartLines"    minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="axId"       type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_BarSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"              type="CT_UnsignedInt"/>
      <xsd:element name="order"            type="CT_UnsignedInt"/>
      <xsd:element name="tx"               type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"             type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="invertIfNegative" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="pictureOptions"   type="CT_PictureOptions"    minOccurs="0"/>
      <xsd:element name="dPt"              type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"            type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="trendline"        type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"          type="CT_ErrBars"           minOccurs="0"/>
      <xsd:element name="cat"              type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"              type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="shape"            type="CT_Shape"             minOccurs="0"/>
      <xsd:element name="extLst"           type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GapAmount">
    <xsd:attribute name="val" type="ST_GapAmount" default="150%"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_GapAmount">
    <xsd:union memberTypes="ST_GapAmountPercent ST_GapAmountUShort"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_GapAmountPercent">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="0*(([0-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)%"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_GapAmountUShort">
    <xsd:restriction base="xsd:unsignedShort">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="500"/>
    </xsd:restriction>
  </xsd:simpleType>
