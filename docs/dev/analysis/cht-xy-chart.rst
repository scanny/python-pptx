.. _XyChart:


X-Y Chart
=========

An X-Y chart, also known as a *scatter* chart, depending on the version of
PowerPoint, is distinguished by having two value axes rather than one
category axis and one value axis.



XML specimen
------------

.. highlight:: xml

XML for default XY chart::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
    <c:chart>
      <c:autoTitleDeleted val="0"/>
      <c:plotArea>
        <c:layout/>
        <c:scatterChart>
          <c:scatterStyle val="lineMarker"/>
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
                    <c:v>Red Bull</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:spPr>
              <a:ln w="47625">
                <a:noFill/>
              </a:ln>
            </c:spPr>
            <c:xVal>
              <c:numRef>
                <c:f>Sheet1!$A$2:$A$7</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="6"/>
                  <c:pt idx="0">
                    <c:v>0.7</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>1.8</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>2.6</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>0.8</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>1.7</c:v>
                  </c:pt>
                  <c:pt idx="5">
                    <c:v>2.5</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:xVal>
            <c:yVal>
              <c:numRef>
                <c:f>Sheet1!$B$2:$B$7</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="6"/>
                  <c:pt idx="0">
                    <c:v>2.7</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>3.2</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>0.8</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:yVal>
            <c:smooth val="0"/>
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
                    <c:v>Monster</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:spPr>
              <a:ln w="47625">
                <a:noFill/>
              </a:ln>
            </c:spPr>
            <c:xVal>
              <c:numRef>
                <c:f>Sheet1!$A$2:$A$7</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="6"/>
                  <c:pt idx="0">
                    <c:v>0.7</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>1.8</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>2.6</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>0.8</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>1.7</c:v>
                  </c:pt>
                  <c:pt idx="5">
                    <c:v>2.5</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:xVal>
            <c:yVal>
              <c:numRef>
                <c:f>Sheet1!$C$2:$C$7</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="6"/>
                  <c:pt idx="3">
                    <c:v>3.2</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>4.3</c:v>
                  </c:pt>
                  <c:pt idx="5">
                    <c:v>1.2</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:yVal>
            <c:smooth val="0"/>
          </c:ser>
          <c:dLbls>
            <c:showLegendKey val="0"/>
            <c:showVal val="0"/>
            <c:showCatName val="0"/>
            <c:showSerName val="0"/>
            <c:showPercent val="0"/>
            <c:showBubbleSize val="0"/>
          </c:dLbls>
          <c:axId val="-2128940872"/>
          <c:axId val="-2129643912"/>
        </c:scatterChart>
        <c:valAx>
          <c:axId val="-2128940872"/>
          <c:scaling>
            <c:orientation val="minMax"/>
          </c:scaling>
          <c:delete val="0"/>
          <c:axPos val="b"/>
          <c:numFmt formatCode="General" sourceLinked="1"/>
          <c:majorTickMark val="out"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="-2129643912"/>
          <c:crosses val="autoZero"/>
          <c:crossBetween val="midCat"/>
        </c:valAx>
        <c:valAx>
          <c:axId val="-2129643912"/>
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
          <c:crossAx val="-2128940872"/>
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
      <c:dispBlanksAs val="gap"/>
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

XY (scatter) chart elements::

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

  <xsd:complexType name="CT_ScatterSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"         type="CT_UnsignedInt"/>
      <xsd:element name="order"       type="CT_UnsignedInt"/>
      <xsd:element name="tx"          type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"        type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="marker"      type="CT_Marker"            minOccurs="0"/>
      <xsd:element name="dPt"         type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"       type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="trendline"   type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"     type="CT_ErrBars"           minOccurs="0" maxOccurs="2"/>
      <xsd:element name="xVal"        type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="yVal"        type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="smooth"      type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_AxDataSource">
    <xsd:sequence>
      <xsd:choice minOccurs="1" maxOccurs="1">
        <xsd:element name="multiLvlStrRef" type="CT_MultiLvlStrRef"/>
        <xsd:element name="numRef"         type="CT_NumRef"/>
        <xsd:element name="numLit"         type="CT_NumData"/>
        <xsd:element name="strRef"         type="CT_StrRef"/>
        <xsd:element name="strLit"         type="CT_StrData"/>
      </xsd:choice>
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

  <xsd:complexType name="CT_MarkerStyle">
    <xsd:attribute name="val" type="ST_MarkerStyle" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_MarkerSize">
    <xsd:attribute name="val" type="ST_MarkerSize" default="5"/>
  </xsd:complexType>

  <xsd:complexType name="CT_NumData">
    <xsd:sequence>
      <xsd:element name="formatCode" type="s:ST_Xstring"     minOccurs="0"/>
      <xsd:element name="ptCount"    type="CT_UnsignedInt"   minOccurs="0"/>
      <xsd:element name="pt"         type="CT_NumVal"        minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_NumDataSource">
    <xsd:sequence>
      <xsd:choice minOccurs="1" maxOccurs="1">
        <xsd:element name="numRef" type="CT_NumRef"/>
        <xsd:element name="numLit" type="CT_NumData"/>
      </xsd:choice>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_NumRef">
    <xsd:sequence>
      <xsd:element name="f"        type="xsd:string"/>
      <xsd:element name="numCache" type="CT_NumData"       minOccurs="0"/>
      <xsd:element name="extLst"   type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:simpleType name="ST_MarkerSize">
    <xsd:restriction base="xsd:unsignedByte">
      <xsd:minInclusive value="2"/>
      <xsd:maxInclusive value="72"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_MarkerStyle">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="circle"/>
      <xsd:enumeration value="dash"/>
      <xsd:enumeration value="diamond"/>
      <xsd:enumeration value="dot"/>
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="picture"/>
      <xsd:enumeration value="plus"/>
      <xsd:enumeration value="square"/>
      <xsd:enumeration value="star"/>
      <xsd:enumeration value="triangle"/>
      <xsd:enumeration value="x"/>
      <xsd:enumeration value="auto"/>
    </xsd:restriction>
  </xsd:simpleType>

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
