
Chart - Plots
=============

Within the ``c:chartSpace/c:chart/c:plotArea`` element may occur 0..n
`xChart` elements, each containing 0..n series.

`xChart` here is a placeholder name for the various plot elements, each of
which ends with `Chart`:

* ``<c:barChart>``
* ``<c:lineChart>``
* ``<c:pieChart>``
* ``<c:areaChart>``
* etc. ...

In the Microsoft API, the term *chart group* is used for the concept these
elements represent. Rather than a group of charts, their role is perhaps
better described as a *series group*. The terminology is a bit vexed when it
comes down to details. The term *plot* was chosen for the purposes of this
library.

The reason the concept of a plot is required is that more than one type of
plotting may appear on a single chart. For example, a line chart can appear
on top of a column chart. To be more precise, one or more series can be
plotted as lines superimposed on one or more series plotted as columns.

The two sets of series both belong to a single chart, and there is only
a single data source to define all the series in the chart. The Excel
workbook that provides the chart data uses only a single worksheet.

Individual series can be assigned a different *chart type* using the
PowerPoint UI; this operation causes them to be placed in a distinct `xChart`
element of the corresponding type. During this operation, PowerPoint adds
`xChart` elements in an order that seems to correspond to logical viewing
order, without respect to the sequence of the series chosen. It appears area
plots are placed in back, bar next, and line in front. This prevents plots of
one type from obscuring the others.

Note that not all combinations of chart types are possible. I've seen area
overlaid with column overlaid with line. Bar cannot be combined with line,
which seems sensible because their axes locations differ (switched
horizontal/vertical).


Feature Summary
---------------

Most plot-level properties are particular to a chart type. One so far is
shared by plots of almost all chart types.

* **Plot.vary_by_categories** -- Read/write boolean determining whether point
  markers for each category have a different color. A point marker here means
  a bar for a bar chart, a pie section, etc. This setting is only operative
  when the plot has a single series. A plot with multiple series uses varying
  colors to distinguish series rather than categories.


Candidate protocol
------------------

Plot.vary_by_categories::

    >>> plot = chart.plots[0]
    >>> plot.vary_by_categories
    True
    >>> plot.vary_by_categories = False
    >>> plot.vary_by_categories
    False


XML Semantics
-------------

* The value of ``c:xChart/c:varyColors`` defaults to |True| when the element
  is not present or when the element is present but its ``val`` attribute is
  not.


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_AreaChart">
    <xsd:sequence>
      <xsd:group    ref="EG_AreaChartShared"/>
      <xsd:element name="axId"   type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Area3DChart">
    <xsd:sequence>
      <xsd:group    ref="EG_AreaChartShared"/>
      <xsd:element name="gapDepth" type="CT_GapAmount"     minOccurs="0"/>
      <xsd:element name="axId"     type="CT_UnsignedInt"   minOccurs="2" maxOccurs="3"/>
      <xsd:element name="extLst"   type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="EG_AreaChartShared">
    <xsd:sequence>
      <xsd:element name="grouping"   type="CT_Grouping"   minOccurs="0"/>
      <xsd:element name="varyColors" type="CT_Boolean"    minOccurs="0"/>
      <xsd:element name="ser"        type="CT_AreaSer"    minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"      type="CT_DLbls"      minOccurs="0"/>
      <xsd:element name="dropLines"  type="CT_ChartLines" minOccurs="0"/>
    </xsd:sequence>
  </xsd:group>

  <xsd:complexType name="CT_Grouping">
    <xsd:attribute name="val" type="ST_Grouping" default="standard"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Grouping">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="percentStacked"/>
      <xsd:enumeration value="standard"/>
      <xsd:enumeration value="stacked"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_BarChart">
    <xsd:sequence>
      <xsd:group    ref="EG_BarChartShared"/>
      <xsd:element name="gapWidth" type="CT_GapAmount"     minOccurs="0"/>
      <xsd:element name="overlap"  type="CT_Overlap"       minOccurs="0"/>
      <xsd:element name="serLines" type="CT_ChartLines"    minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="axId"     type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst"   type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Bar3DChart">
    <xsd:sequence>
      <xsd:group    ref="EG_BarChartShared"/>
      <xsd:element name="gapWidth" type="CT_GapAmount"     minOccurs="0"/>
      <xsd:element name="gapDepth" type="CT_GapAmount"     minOccurs="0"/>
      <xsd:element name="shape"    type="CT_Shape"         minOccurs="0"/>
      <xsd:element name="axId"     type="CT_UnsignedInt"   minOccurs="2" maxOccurs="3"/>
      <xsd:element name="extLst"   type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="EG_BarChartShared">
    <xsd:sequence>
      <xsd:element name="barDir"     type="CT_BarDir"/>
      <xsd:element name="grouping"   type="CT_BarGrouping" minOccurs="0"/>
      <xsd:element name="varyColors" type="CT_Boolean"     minOccurs="0"/>
      <xsd:element name="ser"        type="CT_BarSer"      minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"      type="CT_DLbls"       minOccurs="0"/>
    </xsd:sequence>
  </xsd:group>

  <xsd:complexType name="CT_BarDir">
    <xsd:attribute name="val" type="ST_BarDir" default="col"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_BarDir">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="bar"/>
      <xsd:enumeration value="col"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_BarGrouping">
    <xsd:attribute name="val" type="ST_BarGrouping" default="clustered"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_BarGrouping">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="percentStacked"/>
      <xsd:enumeration value="clustered"/>
      <xsd:enumeration value="standard"/>
      <xsd:enumeration value="stacked"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Shape">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="cone"/>
      <xsd:enumeration value="coneToMax"/>
      <xsd:enumeration value="box"/>
      <xsd:enumeration value="cylinder"/>
      <xsd:enumeration value="pyramid"/>
      <xsd:enumeration value="pyramidToMax"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_LineChart">
    <xsd:sequence>
      <xsd:group ref="EG_LineChartShared" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="hiLowLines" type="CT_ChartLines"    minOccurs="0"/>
      <xsd:element name="upDownBars" type="CT_UpDownBars"    minOccurs="0"/>
      <xsd:element name="marker"     type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="smooth"     type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="axId"       type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="EG_LineChartShared">
    <xsd:sequence>
      <xsd:element name="grouping"   type="CT_Grouping"/>
      <xsd:element name="varyColors" type="CT_Boolean"    minOccurs="0"/>
      <xsd:element name="ser"        type="CT_LineSer"    minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"      type="CT_DLbls"      minOccurs="0"/>
      <xsd:element name="dropLines"  type="CT_ChartLines" minOccurs="0"/>
    </xsd:sequence>
  </xsd:group>

  <xsd:complexType name="CT_PieChart">
    <xsd:sequence>
      <xsd:group    ref="EG_PieChartShared"/>
      <xsd:element name="firstSliceAng" type="CT_FirstSliceAng" minOccurs="0"/>
      <xsd:element name="extLst"        type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="EG_PieChartShared">
    <xsd:sequence>
      <xsd:element name="varyColors" type="CT_Boolean" minOccurs="0"/>
      <xsd:element name="ser"        type="CT_PieSer"  minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"      type="CT_DLbls"   minOccurs="0"/>
    </xsd:sequence>
  </xsd:group>

  <xsd:complexType name="CT_PieSer">
    <xsd:sequence>
      <xsd:group    ref="EG_SerShared"/>
      <xsd:element name="explosion" type="CT_UnsignedInt"   minOccurs="0"/>
      <xsd:element name="dPt"       type="CT_DPt"           minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"     type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="cat"       type="CT_AxDataSource"  minOccurs="0"/>
      <xsd:element name="val"       type="CT_NumDataSource" minOccurs="0"/>
      <xsd:element name="extLst"    type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="EG_SerShared">
    <xsd:sequence>
      <xsd:element name="idx"   type="CT_UnsignedInt"/>
      <xsd:element name="order" type="CT_UnsignedInt"/>
      <xsd:element name="tx"    type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"  type="a:CT_ShapeProperties" minOccurs="0"/>
    </xsd:sequence>
  </xsd:group>

  <xsd:complexType name="CT_RadarChart">
    <xsd:sequence>
      <xsd:element name="radarStyle" type="CT_RadarStyle"/>
      <xsd:element name="varyColors" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="ser"        type="CT_RadarSer"      minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"      type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="axId"       type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_RadarStyle">
    <xsd:attribute name="val" type="ST_RadarStyle" default="standard"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_RadarStyle">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="standard"/>
      <xsd:enumeration value="marker"/>
      <xsd:enumeration value="filled"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_ScatterChart">
    <xsd:sequence>
      <xsd:element name="scatterStyle" type="CT_ScatterStyle"/>
      <xsd:element name="varyColors"   type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="ser"          type="CT_ScatterSer"    minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"        type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="axId"         type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ScatterStyle">
    <xsd:attribute name="val" type="ST_ScatterStyle" default="marker"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_ScatterStyle">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="line"/>
      <xsd:enumeration value="lineMarker"/>
      <xsd:enumeration value="marker"/>
      <xsd:enumeration value="smooth"/>
      <xsd:enumeration value="smoothMarker"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_StockChart">
    <xsd:sequence>
      <xsd:element name="ser"        type="CT_LineSer"       minOccurs="3" maxOccurs="4"/>
      <xsd:element name="dLbls"      type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="dropLines"  type="CT_ChartLines"    minOccurs="0"/>
      <xsd:element name="hiLowLines" type="CT_ChartLines"    minOccurs="0"/>
      <xsd:element name="upDownBars" type="CT_UpDownBars"    minOccurs="0"/>
      <xsd:element name="axId"       type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SurfaceChart">
    <xsd:sequence>
      <xsd:group    ref="EG_SurfaceChartShared"/>
      <xsd:element name="axId"   type="CT_UnsignedInt"   minOccurs="2" maxOccurs="3"/>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Surface3DChart">
    <xsd:sequence>
      <xsd:group ref="EG_SurfaceChartShared"/>
      <xsd:element name="axId"   type="CT_UnsignedInt"   minOccurs="3" maxOccurs="3"/>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="EG_SurfaceChartShared">
    <xsd:sequence>
      <xsd:element name="wireframe" type="CT_Boolean"    minOccurs="0"/>
      <xsd:element name="ser"       type="CT_SurfaceSer" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="bandFmts"  type="CT_BandFmts"   minOccurs="0"/>
    </xsd:sequence>
  </xsd:group>
