.. _RadarChart:


Radar Chart
===========

A radar chart is essentially a line chart wrapped around on itself.

It's worth a try to see if inheriting from LineChart would work.


PowerPoint UI
-------------

* Insert > Chart ... > Other > Radar

* Default Radar uses 5 categories, 2 series, points connected with lines, no
  data point markers. Other radar options can add markers or fill.

* Layout seems an awful lot like a line chart

* Supports data labels, but only with data point value, not custom text (from
  UI).


XML semantics
-------------

* There is one `c:ser` element for each series. Each `c:ser` element contains
  both the categories and the values (even though that is often redundant).

* ? What happens if the category items in a radar chart don't match?
  Inside are `c:xVal`, `c:yVal`, and
  `c:bubbleSize`, each containing the set of points for that "series".


XML specimen
------------

.. highlight:: xml

XML for default radar chart (simplified)::

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
        <c:radarChart>
          <c:radarStyle val="marker"/>
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
            <c:marker>
              <c:symbol val="none"/>
            </c:marker>
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
            <c:marker>
              <c:symbol val="none"/>
            </c:marker>
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
          <c:axId val="2079682968"/>
          <c:axId val="2079686056"/>
        </c:radarChart>
        <c:catAx>
          <c:axId val="2079682968"/>
          <c:scaling>
            <c:orientation val="minMax"/>
          </c:scaling>
          <c:delete val="0"/>
          <c:axPos val="b"/>
          <c:majorGridlines/>
          <c:numFmt formatCode="m/d/yy" sourceLinked="1"/>
          <c:majorTickMark val="out"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="2079686056"/>
          <c:crosses val="autoZero"/>
          <c:auto val="1"/>
          <c:lblAlgn val="ctr"/>
          <c:lblOffset val="100"/>
          <c:noMultiLvlLbl val="0"/>
        </c:catAx>
        <c:valAx>
          <c:axId val="2079686056"/>
          <c:scaling>
            <c:orientation val="minMax"/>
          </c:scaling>
          <c:delete val="0"/>
          <c:axPos val="l"/>
          <c:majorGridlines/>
          <c:numFmt formatCode="General" sourceLinked="1"/>
          <c:majorTickMark val="cross"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="2079682968"/>
          <c:crosses val="autoZero"/>
          <c:crossBetween val="between"/>
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


MS API Protocol
---------------

.. highlight:: vb.net

Make a radar chart::

  Public Sub MakeGraph()

      Dim title       As String = "Title"
      Dim x_label     As String = "X Axis Label"
      Dim y_label     As String = "Y Axis Label"
      Dim new_chart   As Chart

      ' Make the chart.
      Set new_chart = Charts.Add()
      ActiveChart.ChartType = xlRadar
      ActiveChart.SetSourceData _
          Source:=sheet.Range("A1:A" & UBound(values)), _
          PlotBy:=xlColumns

      ' Set the chart's title abd axis labels.
      With ActiveChart
          .HasTitle = True
          .ChartTitle.Characters.Text = title

          .Axes(xlCategory, xlPrimary).HasTitle = True
          .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = x_label

          .Axes(xlValue, xlPrimary).HasTitle = True
          .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = y_label
      End With
  End Sub


Related Schema Definitions
--------------------------

.. highlight:: xml

Radar chart elements::

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

  <xsd:complexType name="CT_RadarSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"    type="CT_UnsignedInt"/>
      <xsd:element name="order"  type="CT_UnsignedInt"/>
      <xsd:element name="tx"     type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="marker" type="CT_Marker"            minOccurs="0"/>
      <xsd:element name="dPt"    type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"  type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="cat"    type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"    type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SerTx">
    <xsd:sequence>
      <xsd:choice minOccurs="1" maxOccurs="1">
        <xsd:element name="strRef" type="CT_StrRef"/>
        <xsd:element name="v"      type="s:ST_Xstring"/>
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

  <xsd:complexType name="CT_DPt">
    <xsd:sequence>
      <xsd:element name="idx"              type="CT_UnsignedInt"/>
      <xsd:element name="invertIfNegative" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="marker"           type="CT_Marker"            minOccurs="0"/>
      <xsd:element name="bubble3D"         type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="explosion"        type="CT_UnsignedInt"       minOccurs="0"/>
      <xsd:element name="spPr"             type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="pictureOptions"   type="CT_PictureOptions"    minOccurs="0"/>
      <xsd:element name="extLst"           type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:simpleType name="ST_RadarStyle">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="standard"/>
      <xsd:enumeration value="marker"/>
      <xsd:enumeration value="filled"/>
    </xsd:restriction>
  </xsd:simpleType>
