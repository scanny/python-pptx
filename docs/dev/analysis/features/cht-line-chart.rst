
Line Chart
==========

A line chart is one of the fundamental plot types.


XML Semantics
-------------

* c:lineChart/c:marker{val=0|1} doesn't seem to have any effect one way or
  the other. Default line charts inserted using PowerPoint always have it set
  to 1 (True). But even if it's set to 0 or removed, markers still appear.
  Hypothesis is that c:lineChart/c:ser/c:marker/c:symbol{val=none} and the
  like are the operative settings.


XML specimens
-------------

.. highlight:: xml

Minimal working XML for a line plot. Note this does not reference values in a
spreadsheet.::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:chart>
      <c:plotArea>
        <c:lineChart>
          <c:grouping val="standard"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx>
              <c:v>Series 1</c:v>
            </c:tx>
            <c:marker>
              <c:symbol val="none"/>
            </c:marker>
            <c:cat>
              <c:strLit>
                <c:ptCount val="2"/>
                <c:pt idx="0">
                  <c:v>Category 1</c:v>
                </c:pt>
                <c:pt idx="1">
                  <c:v>Category 2</c:v>
                </c:pt>
              </c:strLit>
            </c:cat>
            <c:val>
              <c:numLit>
                <c:ptCount val="2"/>
                <c:pt idx="0">
                  <c:v>4.3</c:v>
                </c:pt>
                <c:pt idx="1">
                  <c:v>2.5</c:v>
                </c:pt>
              </c:numLit>
            </c:val>
            <c:smooth val="0"/>
          </c:ser>
          <c:axId val="-2044963928"/>
          <c:axId val="-2071826904"/>
        </c:lineChart>
        <c:catAx>
          <c:axId val="-2044963928"/>
          <c:scaling/>
          <c:delete val="0"/>
          <c:axPos val="b"/>
          <c:crossAx val="-2071826904"/>
        </c:catAx>
        <c:valAx>
          <c:axId val="-2071826904"/>
          <c:scaling/>
          <c:delete val="0"/>
          <c:axPos val="l"/>
          <c:crossAx val="-2044963928"/>
        </c:valAx>
      </c:plotArea>
    </c:chart>
  </c:chartSpace>


Related Schema Definitions
--------------------------

.. highlight:: xml

Line chart elements::

  <xsd:complexType name="CT_LineChart">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="grouping"   type="CT_Grouping"/>
      <xsd:element name="varyColors" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="ser"        type="CT_LineSer"       minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"      type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="dropLines"  type="CT_ChartLines"    minOccurs="0"/>
      <xsd:element name="hiLowLines" type="CT_ChartLines"    minOccurs="0"/>
      <xsd:element name="upDownBars" type="CT_UpDownBars"    minOccurs="0"/>
      <xsd:element name="marker"     type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="smooth"     type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="axId"       type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

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

  <xsd:complexType name="CT_LineSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"       type="CT_UnsignedInt"/>
      <xsd:element name="order"     type="CT_UnsignedInt"/>
      <xsd:element name="tx"        type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"      type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="marker"    type="CT_Marker"            minOccurs="0"/>
      <xsd:element name="dPt"       type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"     type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="trendline" type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"   type="CT_ErrBars"           minOccurs="0"/>
      <xsd:element name="cat"       type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"       type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="smooth"    type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="extLst"    type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Marker">
    <xsd:sequence>
      <xsd:element name="symbol" type="CT_MarkerStyle"       minOccurs="0"/>
      <xsd:element name="size"   type="CT_MarkerSize"        minOccurs="0"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>
