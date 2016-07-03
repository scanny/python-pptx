
Area Chart
==========

The area chart is similar to a stacked line chart where the area between the
lines is filled in.

XML specimens
-------------

.. highlight:: xml

XML for default Area chart::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
    <c:date1904 val="0"/>
    <c:lang val="en-US"/>
    <c:roundedCorners val="0"/>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14">
        <c14:style val="118"/>
      </mc:Choice>
      <mc:Fallback>
        <c:style val="18"/>
      </mc:Fallback>
    </mc:AlternateContent>
    <c:chart>
      <c:autoTitleDeleted val="0"/>
      <c:plotArea>
        <c:layout/>
        <c:areaChart>
          <c:grouping val="standard"/>
          <c:varyColors val="0"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx>
              <c:strRef>
                <c:f>Sheet1!$B$1</c:f>
                <c:strCache>
                  <c:ptCount val="1"/>
                  <c:pt idx="0">
                    <c:v>Series 1</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:cat>
              <c:numRef>
                <c:f>Sheet1!$A$2:$A$6</c:f>
                <c:numCache>
                  <c:formatCode>m/d/yy</c:formatCode>
                  <c:ptCount val="5"/>
                  <c:pt idx="0">
                    <c:v>37261.0</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>37262.0</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>37263.0</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>37264.0</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>37265.0</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:cat>
            <c:val>
              <c:numRef>
                <c:f>Sheet1!$B$2:$B$6</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="5"/>
                  <c:pt idx="0">
                    <c:v>32.0</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>32.0</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>28.0</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>12.0</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>15.0</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:val>
          </c:ser>
          <c:ser>
            <c:idx val="1"/>
            <c:order val="1"/>
            <c:tx>
              <c:strRef>
                <c:f>Sheet1!$C$1</c:f>
                <c:strCache>
                  <c:ptCount val="1"/>
                  <c:pt idx="0">
                    <c:v>Series 2</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:cat>
              <c:numRef>
                <c:f>Sheet1!$A$2:$A$6</c:f>
                <c:numCache>
                  <c:formatCode>m/d/yy</c:formatCode>
                  <c:ptCount val="5"/>
                  <c:pt idx="0">
                    <c:v>37261.0</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>37262.0</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>37263.0</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>37264.0</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>37265.0</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:cat>
            <c:val>
              <c:numRef>
                <c:f>Sheet1!$C$2:$C$6</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="5"/>
                  <c:pt idx="0">
                    <c:v>12.0</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>12.0</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>12.0</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>21.0</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>28.0</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
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
          <c:axId val="-2088330696"/>
          <c:axId val="-2081266648"/>
        </c:areaChart>
        <c:dateAx>
          <c:axId val="-2088330696"/>
          <c:scaling>
            <c:orientation val="minMax"/>
          </c:scaling>
          <c:delete val="0"/>
          <c:axPos val="b"/>
          <c:numFmt formatCode="m/d/yy" sourceLinked="1"/>
          <c:majorTickMark val="out"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="-2081266648"/>
          <c:crosses val="autoZero"/>
          <c:auto val="1"/>
          <c:lblOffset val="100"/>
          <c:baseTimeUnit val="days"/>
        </c:dateAx>
        <c:valAx>
          <c:axId val="-2081266648"/>
          <c:scaling>
            <c:orientation val="minMax"/>
          </c:scaling>
          <c:delete val="0"/>
          <c:axPos val="l"/>
          <c:majorGridlines/>
          <c:numFmt formatCode="General" sourceLinked="1"/>
          <c:majorTickMark val="out"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="-2088330696"/>
          <c:crosses val="autoZero"/>
          <c:crossBetween val="midCat"/>
        </c:valAx>
      </c:plotArea>
      <c:legend>
        <c:legendPos val="r"/>
        <c:layout/>
        <c:overlay val="0"/>
      </c:legend>
      <c:plotVisOnly val="1"/>
      <c:dispBlanksAs val="zero"/>
      <c:showDLblsOverMax val="0"/>
    </c:chart>
    <c:txPr>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
        <a:pPr>
          <a:defRPr sz="1800"/>
        </a:pPr>
        <a:endParaRPr lang="en-US"/>
      </a:p>
    </c:txPr>
    <c:externalData r:id="rId1">
      <c:autoUpdate val="0"/>
    </c:externalData>
  </c:chartSpace>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_AreaChart">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="grouping"   type="CT_Grouping"      minOccurs="0"/>
      <xsd:element name="varyColors" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="ser"        type="CT_AreaSer"       minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"      type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="dropLines"  type="CT_ChartLines"    minOccurs="0"/>
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

  <xsd:complexType name="CT_AreaSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"             type="CT_UnsignedInt"/>
      <xsd:element name="order"           type="CT_UnsignedInt"/>
      <xsd:element name="tx"              type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"            type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="pictureOptions"  type="CT_PictureOptions"    minOccurs="0"/>
      <xsd:element name="dPt"             type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"           type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="trendline"       type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"         type="CT_ErrBars"           minOccurs="0" maxOccurs="2"/>
      <xsd:element name="cat"             type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"             type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="extLst"          type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ChartLines">
    <xsd:sequence>
      <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>
