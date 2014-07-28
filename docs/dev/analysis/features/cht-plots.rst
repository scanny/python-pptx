
Chart - Plots
=============


XML Semantics
-------------

* c:lineChart/c:marker{val=0|1} doesn't seem to have any effect one way or
  the other. Default line charts inserted using PowerPoint always have it set
  to 1 (True). But even if it's set to 0 or removed, markers still appear.
  Hypothesis is that c:lineChart/c:ser/c:marker/c:symbol{val=none} and the
  like are the operative settings.


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
