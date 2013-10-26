############
``CT_Chart``
############

.. highlight:: xml

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 15, 50

   Schema Name  , CT_Chart
   Spec Name    , Chart
   Tag(s)       , c:chart
   Namespace    , drawingml/chart (dml-chart.xsd)
   Schema Line  , 1331
   Spec Section , 21.2.2.27


Attributes
==========

none.


Child elements
==============

================  ===  ================  ==============
name               #   type              line
================  ===  ================  ==============
title              ?   CT_Title          177 dml-chart
autoTitleDeleted   ?   CT_Boolean        17 dml-chart
pivotFmts          ?   CT_PivotFmts      1278 dml-chart
view3D             ?   CT_View3D         236 dml-chart
floor              ?   CT_Surface        247 dml-chart
sideWall           ?   CT_Surface        247 dml-chart
backWall           ?   CT_Surface        247 dml-chart
plotArea           1   CT_PlotArea       1236 dml-chart
legend             ?   CT_Legend         1310 dml-chart
plotVisOnly        ?   CT_Boolean        17 dml-chart
dispBlanksAs       ?   CT_DispBlanksAs   1328 dml-chart
showDLblsOverMax   ?   CT_Boolean        17 dml-chart
extLst             ?   CT_ExtensionList  35 dml-chart
================  ===  ================  ==============


Resources
=========

* ISO-IEC-29500-1, Section 21.2 DrawingML - Charts, pp3354
* ISO-IEC-29500-1, Section 21.2.2.26 chart (Reference to Chart Part)
* ISO-IEC-29500-1, Section 21.2.2.27 chart (Chart), pp3368


Spec text
=========

    This element specifies the chart.


Schema excerpt
==============

::

  <xsd:complexType name="CT_Chart">
    <xsd:sequence>
      <xsd:element name="title"            type="CT_Title"         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="autoTitleDeleted" type="CT_Boolean"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="pivotFmts"        type="CT_PivotFmts"     minOccurs="0" maxOccurs="1"/>
      <xsd:element name="view3D"           type="CT_View3D"        minOccurs="0" maxOccurs="1"/>
      <xsd:element name="floor"            type="CT_Surface"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sideWall"         type="CT_Surface"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="backWall"         type="CT_Surface"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="plotArea"         type="CT_PlotArea"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="legend"           type="CT_Legend"        minOccurs="0" maxOccurs="1"/>
      <xsd:element name="plotVisOnly"      type="CT_Boolean"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="dispBlanksAs"     type="CT_DispBlanksAs"  minOccurs="0" maxOccurs="1"/>
      <xsd:element name="showDLblsOverMax" type="CT_Boolean"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"           type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>

  <xsd:complexType name="CT_Title">
    <xsd:sequence>
      <xsd:element name="tx"      type="CT_Tx"                minOccurs="0" maxOccurs="1"/>
      <xsd:element name="layout"  type="CT_Layout"            minOccurs="0" maxOccurs="1"/>
      <xsd:element name="overlay" type="CT_Boolean"           minOccurs="0" maxOccurs="1"/>
      <xsd:element name="spPr"    type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="txPr"    type="a:CT_TextBody"        minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"  type="CT_ExtensionList"     minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Boolean">
    <xsd:attribute name="val" type="xsd:boolean" use="optional" default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PivotFmts">
    <xsd:sequence>
      <xsd:element name="pivotFmt" type="CT_PivotFmt" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_View3D">
    <xsd:sequence>
      <xsd:element name="rotX"         type="CT_RotX"          minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hPercent"     type="CT_HPercent"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="rotY"         type="CT_RotY"          minOccurs="0" maxOccurs="1"/>
      <xsd:element name="depthPercent" type="CT_DepthPercent"  minOccurs="0" maxOccurs="1"/>
      <xsd:element name="rAngAx"       type="CT_Boolean"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="perspective"  type="CT_Perspective"   minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Surface">
    <xsd:sequence>
      <xsd:element name="thickness" type="CT_Thickness" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="pictureOptions" type="CT_PictureOptions" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_PlotArea">
    <xsd:sequence>
      <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
      <xsd:choice minOccurs="1" maxOccurs="unbounded">
        <xsd:element name="areaChart"      type="CT_AreaChart"      minOccurs="1" maxOccurs="1"/>
        <xsd:element name="area3DChart"    type="CT_Area3DChart"    minOccurs="1" maxOccurs="1"/>
        <xsd:element name="lineChart"      type="CT_LineChart"      minOccurs="1" maxOccurs="1"/>
        <xsd:element name="line3DChart"    type="CT_Line3DChart"    minOccurs="1" maxOccurs="1"/>
        <xsd:element name="stockChart"     type="CT_StockChart"     minOccurs="1" maxOccurs="1"/>
        <xsd:element name="radarChart"     type="CT_RadarChart"     minOccurs="1" maxOccurs="1"/>
        <xsd:element name="scatterChart"   type="CT_ScatterChart"   minOccurs="1" maxOccurs="1"/>
        <xsd:element name="pieChart"       type="CT_PieChart"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="pie3DChart"     type="CT_Pie3DChart"     minOccurs="1" maxOccurs="1"/>
        <xsd:element name="doughnutChart"  type="CT_DoughnutChart"  minOccurs="1" maxOccurs="1"/>
        <xsd:element name="barChart"       type="CT_BarChart"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="bar3DChart"     type="CT_Bar3DChart"     minOccurs="1" maxOccurs="1"/>
        <xsd:element name="ofPieChart"     type="CT_OfPieChart"     minOccurs="1" maxOccurs="1"/>
        <xsd:element name="surfaceChart"   type="CT_SurfaceChart"   minOccurs="1" maxOccurs="1"/>
        <xsd:element name="surface3DChart" type="CT_Surface3DChart" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="bubbleChart"    type="CT_BubbleChart"    minOccurs="1" maxOccurs="1"/>
      </xsd:choice>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="valAx"  type="CT_ValAx"  minOccurs="1" maxOccurs="1"/>
        <xsd:element name="catAx"  type="CT_CatAx"  minOccurs="1" maxOccurs="1"/>
        <xsd:element name="dateAx" type="CT_DateAx" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="serAx"  type="CT_SerAx"  minOccurs="1" maxOccurs="1"/>
      </xsd:choice>
      <xsd:element name="dTable" type="CT_DTable"            minOccurs="0" maxOccurs="1"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst" type="CT_ExtensionList"     minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Legend">
    <xsd:sequence>
      <xsd:element name="legendPos"   type="CT_LegendPos"         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="legendEntry" type="CT_LegendEntry"       minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="layout"      type="CT_Layout"            minOccurs="0" maxOccurs="1"/>
      <xsd:element name="overlay"     type="CT_Boolean"           minOccurs="0" maxOccurs="1"/>
      <xsd:element name="spPr"        type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="txPr"        type="a:CT_TextBody"        minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"     minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_DispBlanksAs">
    <xsd:attribute name="val" type="ST_DispBlanksAs" default="zero"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ExtensionList">
    <xsd:sequence>
      <xsd:element name="ext" type="CT_Extension" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>
