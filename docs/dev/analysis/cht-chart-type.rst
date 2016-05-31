
Chart - Chart Type
==================


c:areaChart
-----------

* ./c:grouping{val=stacked}        => AREA_STACKED
* ./c:grouping{val=percentStacked} => AREA_STACKED_100
* ./c:grouping{val=standard}       => AREA
* .                                => AREA


c:area3DChart
-------------

* ./c:grouping{val=stacked}        => THREE_D_AREA_STACKED
* ./c:grouping{val=percentStacked} => THREE_D_AREA_STACKED_100
* ./c:grouping{val=standard}       => THREE_D_AREA
* .                                => THREE_D_AREA


c:barChart
----------

* ./c:barDir{val=bar}

  + ./c:grouping{val=clustered}      => BAR_CLUSTERED
  + ./c:grouping{val=stacked}        => BAR_STACKED
  + ./c:grouping{val=percentStacked} => BAR_STACKED_100

* ./c:barDir{val=col}

  + ./c:grouping{val=clustered}      => COLUMN_CLUSTERED
  + ./c:grouping{val=stacked}        => COLUMN_STACKED
  + ./c:grouping{val=percentStacked} => COLUMN_STACKED_100


c:bar3DChart
------------

* ./c:barDir{val=bar}

  + ./c:grouping{val=clustered}

    - ./c:shape{val=box}      => THREE_D_BAR_CLUSTERED
    - ./c:shape{val=cone}     => CONE_BAR_CLUSTERED
    - ./c:shape{val=cylinder} => CYLINDER_BAR_CLUSTERED
    - ./c:shape{val=pyramid}  => PYRAMID_BAR_CLUSTERED

  + ./c:grouping{val=stacked}

    - ./c:shape{val=box}      => THREE_D_BAR_STACKED
    - ./c:shape{val=cone}     => CONE_BAR_STACKED
    - ./c:shape{val=cylinder} => CYLINDER_BAR_STACKED
    - ./c:shape{val=pyramid}  => PYRAMID_BAR_STACKED

  + ./c:grouping{val=percentStacked}

    - ./c:shape{val=box}      => THREE_D_BAR_STACKED_100
    - ./c:shape{val=cone}     => CONE_BAR_STACKED_100
    - ./c:shape{val=cylinder} => CYLINDER_BAR_STACKED_100
    - ./c:shape{val=pyramid}  => PYRAMID_BAR_STACKED_100

* ./c:barDir{val=col}

  + ./c:grouping{val=clustered}

    - ./c:shape{val=box}      => THREE_D_COLUMN_CLUSTERED
    - ./c:shape{val=cone}     => CONE_COL_CLUSTERED
    - ./c:shape{val=cylinder} => CYLINDER_COL_CLUSTERED
    - ./c:shape{val=pyramid}  => PYRAMID_COL_CLUSTERED

  + ./c:grouping{val=stacked}

    - ./c:shape{val=box}      => THREE_D_COLUMN_STACKED
    - ./c:shape{val=cone}     => CONE_COL_STACKED
    - ./c:shape{val=cylinder} => CYLINDER_COL_STACKED
    - ./c:shape{val=pyramid}  => PYRAMID_COL_STACKED

  + ./c:grouping{val=percentStacked}

    - ./c:shape{val=box}      => THREE_D_COLUMN_STACKED_100
    - ./c:shape{val=cone}     => CONE_COL_STACKED_100
    - ./c:shape{val=cylinder} => CYLINDER_COL_STACKED_100
    - ./c:shape{val=pyramid}  => PYRAMID_COL_STACKED_100

  + ./c:grouping{val=standard}

    - ./c:shape{val=box}      => THREE_D_COLUMN
    - ./c:shape{val=cone}     => CONE_COL
    - ./c:shape{val=cylinder} => CYLINDER_COL
    - ./c:shape{val=pyramid}  => PYRAMID_COL


c:bubbleChart
-------------

* ./c:bubble3D{val=0} => BUBBLE
* ./c:bubble3D{val=1} => BUBBLE_THREE_D_EFFECT


c:doughnutChart
---------------

* .                          => DOUGHNUT
* ./c:ser/c:explosion{val>0} => DOUGHNUT_EXPLODED


c:lineChart
-----------

* ./c:grouping{val=standard}

  + ./c:ser/c:marker/c:symbol{val=none} => LINE
  + ./c:marker{val=1}                   => LINE_MARKERS

* ./c:grouping{val=stacked}

  + ./c:marker{val=1}                   => LINE_MARKERS_STACKED
  + ./c:ser/c:marker/c:symbol{val=none} => LINE_STACKED

* ./c:grouping{val=percentStacked}

  + ./c:marker{val=1}                   => LINE_MARKERS_STACKED_100
  + ./c:ser/c:marker/c:symbol{val=none} => LINE_STACKED_100


c:line3DChart
-------------

* . => THREE_D_LINE


c:ofPieChart
------------

* ./c:ofPieType{val=bar} => BAR_OF_PIE
* ./c:ofPieType{val=pie} => PIE_OF_PIE


c:pieChart
----------

* .                          => PIE
* ./c:ser/c:explosion{val>0} => PIE_EXPLODED


c:pie3DChart
------------

* .                          => THREE_D_PIE
* ./c:ser/c:explosion{val>0} => THREE_D_PIE_EXPLODED


c:radarChart
------------

* ./c:radarStyle{val=standard} => RADAR
* ./c:radarStyle{val=filled}   => RADAR_FILLED
* ./c:radarStyle{val=marker}   => RADAR_MARKERS


c:scatterChart
--------------

* ./c:scatterStyle{val=lineMarker}   =>  XY_SCATTER
* has to do with ./c:ser/c:spPr/a:ln/a:noFill
* ./c:scatterStyle{val=lineMarker}   =>  XY_SCATTER_LINES
* ./c:scatterStyle{val=line}         =>  XY_SCATTER_LINES_NO_MARKERS
* ./c:scatterStyle{val=smoothMarker} =>  XY_SCATTER_SMOOTH
* ./c:scatterStyle{val=smooth}       =>  XY_SCATTER_SMOOTH_NO_MARKERS
* check all these to verify


c:stockChart
------------

* ./? => STOCK_HLC
* ./? => STOCK_OHLC
* ./? => STOCK_VHLC
* ./? => STOCK_VOHLC
* possibly related to 3 vs. 4 series. VOHLC has a second plot and axis for
  volume


c:surface3DChart
----------------

* ./c:wireframe{val=0} => SURFACE_TOP_VIEW
* ./c:wireframe{val=1} => SURFACE_TOP_VIEW_WIREFRAME


c:surfaceChart
--------------

* ./c:wireframe{val=0} => SURFACE
* ./c:wireframe{val=1} => SURFACE_WIREFRAME


XML specimens
-------------

.. highlight:: xml

::

  <c:barChart>
    <c:barDir val="col"/>
    <c:grouping val="clustered"/>
    <c:ser>
      <c:idx val="0"/>
      <c:order val="0"/>
      <c:cat>...</c:cat>
      <c:val>...</c:val>
    </c:ser>
    <c:axId val="-2068027336"/>
    <c:axId val="-2113994440"/>
  </c:barChart>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

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
