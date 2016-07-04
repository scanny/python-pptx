
Charts - Overview
=================

Adding a new chart - steps and tests
------------------------------------

1. Analysis - add chart type page with schema for each new chart element and
   see if there's anything distinctive about the chart type series

2. feature/sld-add-chart .. add new chart types, steps/chart.py

3. pptx/chart/xmlwriter.py add new type to ChartXmlWriter

4. pptx/chart/xmlwriter.py add _AreaChartXmlWriter, one per new element type

5. pptx/chart/series.py add AreaSeries, one per new element type


There are 73 different possible chart types, but only 16 distinct XML
chart-type elements. The implementation effort is largely proportional to the
number of new XML chart-type elements.

Here is an accounting of the implementation status of the 73 chart types:


29 - supported for creation so far
++++++++++++++++++++++++++++++++++

areaChart
~~~~~~~~~

* AREA
* AREA_STACKED
* AREA_STACKED_100

barChart
~~~~~~~~

* BAR_CLUSTERED
* BAR_STACKED
* BAR_STACKED_100
* COLUMN_CLUSTERED
* COLUMN_STACKED
* COLUMN_STACKED_100

bubbleChart
~~~~~~~~~~~

* BUBBLE
* BUBBLE_THREE_D_EFFECT

doughnutChart
~~~~~~~~~~~~~

* DOUGHNUT
* DOUGHNUT_EXPLODED

lineChart
~~~~~~~~~

* LINE
* LINE_MARKERS
* LINE_MARKERS_STACKED
* LINE_MARKERS_STACKED_100
* LINE_STACKED
* LINE_STACKED_100

pieChart
~~~~~~~~

* PIE
* PIE_EXPLODED

radarChart
~~~~~~~~~~

* RADAR
* RADAR_FILLED
* RADAR_MARKERS

scatterChart
~~~~~~~~~~~~

* XY_SCATTER
* XY_SCATTER_LINES
* XY_SCATTER_LINES_NO_MARKERS
* XY_SCATTER_SMOOTH
* XY_SCATTER_SMOOTH_NO_MARKERS

44 remaining:
+++++++++++++

area3DChart
~~~~~~~~~~~

* THREE_D_AREA
* THREE_D_AREA_STACKED
* THREE_D_AREA_STACKED_100

bar3DChart (28 types)
~~~~~~~~~~~~~~~~~~~~~

* THREE_D_BAR_CLUSTERED
* THREE_D_BAR_STACKED
* THREE_D_BAR_STACKED_100
* THREE_D_COLUMN
* THREE_D_COLUMN_CLUSTERED
* THREE_D_COLUMN_STACKED
* THREE_D_COLUMN_STACKED_100

* CONE_BAR_CLUSTERED
* CONE_BAR_STACKED
* CONE_BAR_STACKED_100
* CONE_COL
* CONE_COL_CLUSTERED
* CONE_COL_STACKED
* CONE_COL_STACKED_100

* CYLINDER_BAR_CLUSTERED
* CYLINDER_BAR_STACKED
* CYLINDER_BAR_STACKED_100
* CYLINDER_COL
* CYLINDER_COL_CLUSTERED
* CYLINDER_COL_STACKED
* CYLINDER_COL_STACKED_100

* PYRAMID_BAR_CLUSTERED
* PYRAMID_BAR_STACKED
* PYRAMID_BAR_STACKED_100
* PYRAMID_COL
* PYRAMID_COL_CLUSTERED
* PYRAMID_COL_STACKED
* PYRAMID_COL_STACKED_100

line3DChart
~~~~~~~~~~~

* THREE_D_LINE

pie3DChart
~~~~~~~~~~

* THREE_D_PIE
* THREE_D_PIE_EXPLODED

ofPieChart
~~~~~~~~~~

* BAR_OF_PIE
* PIE_OF_PIE

stockChart
~~~~~~~~~~

* STOCK_HLC
* STOCK_OHLC
* STOCK_VHLC
* STOCK_VOHLC

surfaceChart
~~~~~~~~~~~~

* SURFACE
* SURFACE_WIREFRAME

surface3DChart
~~~~~~~~~~~~~~

* SURFACE_TOP_VIEW
* SURFACE_TOP_VIEW_WIREFRAME


Chart parts glossary
--------------------

**data point (point)**
   An individual numeric value, represented by a bar, point, column, or pie
   slice.

**data series (series)**
   A group of related data points. For example, the columns of a series will
   all be the same color.

**category axis (X axis)**
   The horizontal axis of a two-dimensional or three-dimensional chart.

**value axis (Y axis)**
   The vertical axis of a two-dimensional or three-dimensional chart.

**depth axis (Z axis)**
   The front-to-back axis of a three-dimensional chart.

**grid lines**
   Horizontal or vertical lines that may be added to an axis to aid
   comparison of a data point to an axis value.

**legend**
   A key that explains which data series each color or pattern represents.

**floor**
   The bottom of a three-dimensional chart.

**walls**
   The background of a chart. Three-dimensional charts have a back wall and
   a side wall, which can be formatted separately.

**data labels**
   Numeric labels on each data point. A data label can represent the actual
   value or a percentage.

**axis title**
   Explanatory text label associated with an axis

**data table**
   A optional tabular display within the *plot area* of the values on which
   the chart is based. Not to be confused with the Excel worksheet holding
   the chart values.

**chart title**
   A label explaining the overall purpose of the chart.

**chart area**
   Overall chart object, containing the chart and all its auxiliary pieces
   such as legends and titles.

**plot area**
   Region of the chart area that contains the actual plots, bounded by but
   not including the axes. May contain more than one plot, each with its own
   distinct set of series. A plot is known as a *chart group* in the MS API.

**axis**
   ... may be either a *category axis* or a *value axis* ... on
   a two-dimensional chart, either the horizontal (*x*) axis or the vertical
   (*y*) axis. A 3-dimensional chart also has a depth (*z*) axis. Pie,
   doughnut, and radar charts have a radial axis.

   How many axes do each of the different chart types have?

**series categories**
   ...

**series values**
   ...


Chart types
-----------

* column

  - 2-D column

    + clustered column
    + stacked column
    + 100% stacked column

  - 3-D column

    + 3-D clustered column
    + 3-D stacked column
    + 3-D 100% stacked column
    + 3-D column

  - cylinder
  - cone
  - pyramid

* line

  + 2-D line
  + 3-D line

* pie

  + 2-D pie
  + 3-D pie

* bar

  + 2-D bar
  + 3-D bar
  + cylinder
  + cone
  + pyramid

* area

* scatter

* other

  + stock (e.g. open-high-low-close)
  + surface
  + doughnut
  + bubble
  + radar


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <!-- homonym <c:chart> element in graphicData element -->

  <xsd:element name="chart" type="CT_RelId"/>

  <xsd:complexType name="CT_RelId">
    <xsd:attribute ref="r:id" use="required"/>
  </xsd:complexType>


  <!-- elements in chartX.xml part -->

  <xsd:element name="chartSpace" type="CT_ChartSpace"/>

  <xsd:complexType name="CT_ChartSpace">
    <xsd:sequence>
      <xsd:element name="date1904"       type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="lang"           type="CT_TextLanguageID"    minOccurs="0"/>
      <xsd:element name="roundedCorners" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="style"          type="CT_Style"             minOccurs="0"/>
      <xsd:element name="clrMapOvr"      type="a:CT_ColorMapping"    minOccurs="0"/>
      <xsd:element name="pivotSource"    type="CT_PivotSource"       minOccurs="0"/>
      <xsd:element name="protection"     type="CT_Protection"        minOccurs="0"/>
      <xsd:element name="chart"          type="CT_Chart"/>
      <xsd:element name="spPr"           type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"           type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="externalData"   type="CT_ExternalData"      minOccurs="0"/>
      <xsd:element name="printSettings"  type="CT_PrintSettings"     minOccurs="0"/>
      <xsd:element name="userShapes"     type="CT_RelId"             minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Chart">
    <xsd:sequence>
      <xsd:element name="title"            type="CT_Title"         minOccurs="0"/>
      <xsd:element name="autoTitleDeleted" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="pivotFmts"        type="CT_PivotFmts"     minOccurs="0"/>
      <xsd:element name="view3D"           type="CT_View3D"        minOccurs="0"/>
      <xsd:element name="floor"            type="CT_Surface"       minOccurs="0"/>
      <xsd:element name="sideWall"         type="CT_Surface"       minOccurs="0"/>
      <xsd:element name="backWall"         type="CT_Surface"       minOccurs="0"/>
      <xsd:element name="plotArea"         type="CT_PlotArea"/>
      <xsd:element name="legend"           type="CT_Legend"        minOccurs="0"/>
      <xsd:element name="plotVisOnly"      type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="dispBlanksAs"     type="CT_DispBlanksAs"  minOccurs="0"/>
      <xsd:element name="showDLblsOverMax" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="extLst"           type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

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

  <xsd:complexType name="CT_Boolean">
    <xsd:attribute name="val" type="xsd:boolean" use="optional" default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Double">
    <xsd:attribute name="val" type="xsd:double" use="required"/>
  </xsd:complexType>

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
    <xsd:attribute name="formatCode"   type="s:ST_Xstring" use="required"/>
    <xsd:attribute name="sourceLinked" type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TickMark">
    <xsd:attribute name="val" type="ST_TickMark" default="cross"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_TickMark">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="cross"/>
      <xsd:enumeration value="in"/>
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="out"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_DLbls">
    <xsd:sequence>
      <xsd:element name="dLbl" type="CT_DLbl" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:choice>
        <xsd:element name="delete" type="CT_Boolean"/>
        <xsd:group   ref="Group_DLbls"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_DLbl">
    <xsd:sequence>
      <xsd:element name="idx" type="CT_UnsignedInt"/>
      <xsd:choice>
        <xsd:element name="delete" type="CT_Boolean"/>
        <xsd:group   ref="Group_DLbl"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="Group_DLbls">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="numFmt"          type="CT_NumFmt"            minOccurs="0"/>
      <xsd:element name="spPr"            type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"            type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="dLblPos"         type="CT_DLblPos"           minOccurs="0"/>
      <xsd:element name="showLegendKey"   type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showVal"         type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showCatName"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showSerName"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showPercent"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showBubbleSize"  type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="separator"       type="xsd:string"           minOccurs="0"/>
      <xsd:element name="showLeaderLines" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="leaderLines"     type="CT_ChartLines"        minOccurs="0"/>
    </xsd:sequence>
  </xsd:group>

  <xsd:complexType name="CT_Style">
    <xsd:attribute name="val" type="ST_Style" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Style">
    <xsd:restriction base="xsd:unsignedByte">
      <xsd:minInclusive value="1"/>
      <xsd:maxInclusive value="48"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_Legend">
    <xsd:sequence>
      <xsd:element name="legendPos"   type="CT_LegendPos"         minOccurs="0"/>
      <xsd:element name="legendEntry" type="CT_LegendEntry"       minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="layout"      type="CT_Layout"            minOccurs="0"/>
      <xsd:element name="overlay"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="spPr"        type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"        type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_LegendPos">
    <xsd:attribute name="val" type="ST_LegendPos" default="r"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_LegendPos">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="b"/>
      <xsd:enumeration value="tr"/>
      <xsd:enumeration value="l"/>
      <xsd:enumeration value="r"/>
      <xsd:enumeration value="t"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_DLblPos">
    <xsd:attribute name="val" type="ST_DLblPos" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_DLblPos">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="bestFit"/>
      <xsd:enumeration value="b"/>
      <xsd:enumeration value="ctr"/>
      <xsd:enumeration value="inBase"/>
      <xsd:enumeration value="inEnd"/>
      <xsd:enumeration value="l"/>
      <xsd:enumeration value="outEnd"/>
      <xsd:enumeration value="r"/>
      <xsd:enumeration value="t"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_LblOffset">
    <xsd:attribute name="val" type="ST_LblOffset" default="100%"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_LblOffset">
    <xsd:union memberTypes="ST_LblOffsetPercent ST_LblOffsetUShort"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_LblOffsetUShort">
    <xsd:restriction base="xsd:unsignedShort">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="1000"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_LblOffsetPercent">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="0*(([0-9])|([1-9][0-9])|([1-9][0-9][0-9])|1000)%"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_Overlap">
    <xsd:attribute name="val" type="ST_Overlap" default="0%"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Overlap">
    <xsd:union memberTypes="ST_OverlapPercent ST_OverlapByte"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_OverlapPercent">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="(-?0*(([0-9])|([1-9][0-9])|100))%"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_OverlapByte">
    <xsd:restriction base="xsd:byte">
      <xsd:minInclusive value="-100"/>
      <xsd:maxInclusive value="100"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_Layout">
    <xsd:sequence>
      <xsd:element name="manualLayout" type="CT_ManualLayout"  minOccurs="0"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ManualLayout">
    <xsd:sequence>
      <xsd:element name="layoutTarget" type="CT_LayoutTarget"  minOccurs="0"/>
      <xsd:element name="xMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="yMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="wMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="hMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="x"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="y"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="w"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="h"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_LayoutMode">
    <xsd:attribute name="val" type="ST_LayoutMode" default="factor"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_LayoutMode">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="edge"/>
      <xsd:enumeration value="factor"/>
    </xsd:restriction>
  </xsd:simpleType>
