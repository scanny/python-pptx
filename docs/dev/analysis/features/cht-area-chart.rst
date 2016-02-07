
Area Chart
==========

The area chart is a type of line chart which displays colored regions under
each line. It includes standard, stacked and stacked 100% varieties by
specifying a grouping.


XML specimens
-------------

.. highlight:: xml

Minimal working XML for a single-series area plot. Note this does not
reference values in a spreadsheet.::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:chart>
      <c:plotArea>
        <c:areaChart>
          <c:grouping val="standard"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:cat>
              <c:strLit>
                <c:ptCount val="2"/>
                <c:pt idx="0">
                  <c:v>Jan</c:v>
                </c:pt>
                <c:pt idx="1">
                  <c:v>Feb</c:v>
                </c:pt>
              </c:strLit>
            </c:cat>
            <c:val>
              <c:numLit>
                <c:ptCount val="2"/>
                <c:pt idx="0">
                  <c:v>75</c:v>
                </c:pt>
                <c:pt idx="1">
                  <c:v>27</c:v>
                </c:pt>
              </c:numLit>
            </c:val>
          </c:ser>
          <c:axId val="-2068027336"/>
          <c:axId val="-2113994440"/>
        </c:areaChart>
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

  <xsd:complexType name="CT_AreaChart">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="grouping"   type="CT_Grouping"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="varyColors" type="CT_Boolean"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="ser"        type="CT_AreaSer"       minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"      type="CT_DLbls"         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="dropLines"  type="CT_ChartLines"    minOccurs="0" maxOccurs="1"/>
      <xsd:element name="axId"       type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_AreaSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"            type="CT_UnsignedInt"       minOccurs="1" maxOccurs="1"/>
      <xsd:element name="order"          type="CT_UnsignedInt"       minOccurs="1" maxOccurs="1"/>
      <xsd:element name="tx"             type="CT_SerTx"             minOccurs="0" maxOccurs="1"/>
      <xsd:element name="spPr"           type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="pictureOptions" type="CT_PictureOptions"    minOccurs="0" maxOccurs="1"/>
      <xsd:element name="dPt"            type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"          type="CT_DLbls"             minOccurs="0" maxOccurs="1"/>
      <xsd:element name="trendline"      type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"        type="CT_ErrBars"           minOccurs="0" maxOccurs="2"/>
      <xsd:element name="cat"            type="CT_AxDataSource"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="val"            type="CT_NumDataSource"     minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <!-- grouping -->

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
