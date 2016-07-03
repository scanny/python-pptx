
Doughnut Chart
==============

A doughnut chart is similar in many ways to a pie chart, except there is
a "hole" in the middle. It can accept multiple series, which appear as
concentric "rings".


XML specimens
-------------

.. highlight:: xml

XML for default doughnut chart::

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
        <c:doughnutChart>
          <c:varyColors val="1"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx>
              <c:strRef>
                <c:f>Sheet1!$B$1</c:f>
                <c:strCache>
                  <c:ptCount val="1"/>
                  <c:pt idx="0">
                    <c:v>Sales</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:cat>
              <c:strRef>
                <c:f>Sheet1!$A$2:$A$5</c:f>
                <c:strCache>
                  <c:ptCount val="4"/>
                  <c:pt idx="0">
                    <c:v>1st Qtr</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>2nd Qtr</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>3rd Qtr</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>4th Qtr</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:cat>
            <c:val>
              <c:numRef>
                <c:f>Sheet1!$B$2:$B$5</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="4"/>
                  <c:pt idx="0">
                    <c:v>8.2</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>3.2</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>1.4</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>1.2</c:v>
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
                    <c:v>Revenue</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:cat>
              <c:strRef>
                <c:f>Sheet1!$A$2:$A$5</c:f>
                <c:strCache>
                  <c:ptCount val="4"/>
                  <c:pt idx="0">
                    <c:v>1st Qtr</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>2nd Qtr</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>3rd Qtr</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>4th Qtr</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:cat>
            <c:val>
              <c:numRef>
                <c:f>Sheet1!$C$2:$C$5</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="4"/>
                  <c:pt idx="0">
                    <c:v>1000.0</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>2000.0</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>3000.0</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>4000.0</c:v>
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
            <c:showLeaderLines val="1"/>
          </c:dLbls>
          <c:firstSliceAng val="0"/>
          <c:holeSize val="50"/>
        </c:doughnutChart>
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

::

  <xsd:complexType name="CT_DoughnutChart">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="varyColors"    type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="ser"           type="CT_PieSer"        minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"         type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="firstSliceAng" type="CT_FirstSliceAng" minOccurs="0"/>
      <xsd:element name="holeSize"      type="CT_HoleSize"      minOccurs="0"/>
      <xsd:element name="extLst"        type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_PieSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"       type="CT_UnsignedInt"/>
      <xsd:element name="order"     type="CT_UnsignedInt"/>
      <xsd:element name="tx"        type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"      type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="explosion" type="CT_UnsignedInt"       minOccurs="0"/>
      <xsd:element name="dPt"       type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"     type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="cat"       type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"       type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="extLst"    type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>
