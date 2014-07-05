
Chart Axes
==========

PowerPoint chart axes come in three varieties: category axis, value axis, and
series axis. A series axis only appears on a 3D chart and is also known as
its depth axis.

A chart may have two category axes and/or two value axes. The second axis if
there is one is known as the *secondary category axis* or *secondary value
axis*.

A category axis may appear as either the horizontal or vertical axis,
depending upon the chart type. Likewise for a value axis.


PowerPoint behavior
-------------------

``<c:delete val="0"/>`` element
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

* when delete element is absent, the default value is True
* when ``val`` attribute is absent, the default value is True


XML specimens
-------------

.. highlight:: xml

Example axes XML for a single-series column plot::

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


Related Schema Definitions
--------------------------

::

  <xsd:complexType name="CT_PlotArea">
    <xsd:sequence>
      <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
      <xsd:choice minOccurs="1" maxOccurs="unbounded">
        <xsd:element name="areaChart"      type="CT_AreaChart"/>
        <xsd:element name="area3DChart"    type="CT_Area3DChart"/>
        <xsd:element name="lineChart"      type="CT_LineChart"/>
        <xsd:element name="line3DChart"    type="CT_Line3DChart"/>
        <xsd:element name="stockChart"     type="CT_StockChart"/>
        <xsd:element name="radarChart"     type="CT_RadarChart"/>
        <xsd:element name="scatterChart"   type="CT_ScatterChart"/>
        <xsd:element name="pieChart"       type="CT_PieChart"/>
        <xsd:element name="pie3DChart"     type="CT_Pie3DChart"/>
        <xsd:element name="doughnutChart"  type="CT_DoughnutChart"/>
        <xsd:element name="barChart"       type="CT_BarChart"/>
        <xsd:element name="bar3DChart"     type="CT_Bar3DChart"/>
        <xsd:element name="ofPieChart"     type="CT_OfPieChart"/>
        <xsd:element name="surfaceChart"   type="CT_SurfaceChart"/>
        <xsd:element name="surface3DChart" type="CT_Surface3DChart"/>
        <xsd:element name="bubbleChart"    type="CT_BubbleChart"/>
      </xsd:choice>
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

  <xsd:complexType name="CT_Boolean">
    <xsd:attribute name="val" type="xsd:boolean" default="true"/>
  </xsd:complexType>
