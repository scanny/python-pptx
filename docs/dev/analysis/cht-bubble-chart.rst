.. _BubbleChart:


Bubble Chart
============

A bubble chart is an extension of an X-Y or scatter chart. Containing an
extra (third) series, the size of the bubble is proportional to the value in
the third series.

The bubble is proportioned to its value by one of two methods selected by the
client:

* Bubble *width* is proportional to value (linear)
* Bubble *area* is proportional to value (quadratic)

By default, the scale of the bubbles is determined by setting the largest
bubble equal to a "default bubble size" equal to roughly 25% of the height or
width of the chart area, whichever is less. The default bubble size can be
scaled using a property setting to obtain proportionally larger or smaller
bubbles.


PowerPoint UI
-------------

To create a bubble chart by hand in PowerPoint:

1. Carts ribbon > Other > Bubble

A three-row, three-column Excel worksheet opens and the default chart appears.


XML semantics
-------------

* I think this is going to get into blank cells if multiple series are
  required. Might be worth checking out how using NA() possibly differs as to
  how it appears in the XML.

* There is a single `c:ser` element. Inside are `c:xVal`, `c:yVal`, and
  `c:bubbleSize`, each containing the set of points for that "series".


XML specimen
------------

.. highlight:: xml

XML for default bubble chart::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
    <c:chart>
      <c:plotArea>
        <c:bubbleChart>
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
                    <c:v>Y-Value 1</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:invertIfNegative val="0"/>
            <c:xVal>
              <c:numRef>
                <c:f>Sheet1!$A$2:$A$4</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="3"/>
                  <c:pt idx="0">
                    <c:v>0.7</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>1.8</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>2.6</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:xVal>
            <c:yVal>
              <c:numRef>
                <c:f>Sheet1!$B$2:$B$4</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="3"/>
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
            <c:bubbleSize>
              <c:numRef>
                <c:f>Sheet1!$C$2:$C$4</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="3"/>
                  <c:pt idx="0">
                    <c:v>10.0</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>4.0</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>8.0</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:bubbleSize>
            <c:bubble3D val="0"/>
          </c:ser>
          <c:dLbls>
            <c:showLegendKey val="0"/>
            <c:showVal val="0"/>
            <c:showCatName val="0"/>
            <c:showSerName val="0"/>
            <c:showPercent val="0"/>
            <c:showBubbleSize val="0"/>
          </c:dLbls>
          <c:bubbleScale val="100"/>
          <c:showNegBubbles val="0"/>
          <c:axId val="2110171512"/>
          <c:axId val="2110299944"/>
        </c:bubbleChart>
        <c:valAx>
          <c:axId val="2110171512"/>
          <c:scaling>
            <c:orientation val="minMax"/>
          </c:scaling>
          <c:delete val="0"/>
          <c:axPos val="b"/>
          <c:numFmt formatCode="General" sourceLinked="1"/>
          <c:majorTickMark val="out"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="2110299944"/>
          <c:crosses val="autoZero"/>
          <c:crossBetween val="midCat"/>
        </c:valAx>
        <c:valAx>
          <c:axId val="2110299944"/>
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
          <c:crossAx val="2110171512"/>
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


MS API Protocol
---------------

.. highlight:: vb.net

Create (unconventional) multi-series bubble chart in Excel::

    If (selection.Columns.Count <> 4 Or selection.Rows.Count < 3) Then
        MsgBox "Selection must have 4 columns and at least 2 rows"
        Exit Sub
    End If

    Dim bubbleChart As ChartObject
    Set bubbleChart = ActiveSheet.ChartObjects.Add(
          Left:=selection.Left, Width:=600, Top:=selection.Top, Height:=400
    )
    bubbleChart.chart.ChartType = xlBubble
    Dim r As Integer
    For r = 2 To selection.Rows.Count
        With bubbleChart.chart.SeriesCollection.NewSeries
            .Name = "=" & selection.Cells(r, 1).Address(External:=True)
            .XValues = selection.Cells(r, 2).Address(External:=True)
            .Values = selection.Cells(r, 3).Address(External:=True)
            .BubbleSizes = selection.Cells(r, 4).Address(External:=True)
        End With
    Next

    bubbleChart.chart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    bubbleChart.chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "=" & selection.Cells(1, 2).Address(External:=True)

    bubbleChart.chart.SetElement (msoElementPrimaryValueAxisTitleRotated)
    bubbleChart.chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "=" & selection.Cells(1, 3).Address(External:=True)

    bubbleChart.chart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    bubbleChart.chart.Axes(xlCategory).MinimumScale = 0


Related Schema Definitions
--------------------------

.. highlight:: xml

Bubble chart elements::

  <xsd:complexType name="CT_BubbleChart">
    <xsd:sequence>
      <xsd:element name="varyColors"     type="CT_Boolean"        minOccurs="0"/>
      <xsd:element name="ser"            type="CT_BubbleSer"      minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"          type="CT_DLbls"          minOccurs="0"/>
      <xsd:element name="bubble3D"       type="CT_Boolean"        minOccurs="0"/>
      <xsd:element name="bubbleScale"    type="CT_BubbleScale"    minOccurs="0"/>
      <xsd:element name="showNegBubbles" type="CT_Boolean"        minOccurs="0"/>
      <xsd:element name="sizeRepresents" type="CT_SizeRepresents" minOccurs="0"/>
      <xsd:element name="axId"           type="CT_UnsignedInt"    minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"  minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_BubbleSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"              type="CT_UnsignedInt"/>
      <xsd:element name="order"            type="CT_UnsignedInt"/>
      <xsd:element name="tx"               type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"             type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="invertIfNegative" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="dPt"              type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"            type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="trendline"        type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"          type="CT_ErrBars"           minOccurs="0" maxOccurs="2"/>
      <xsd:element name="xVal"             type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="yVal"             type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="bubbleSize"       type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="bubble3D"         type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="extLst"           type="CT_ExtensionList"     minOccurs="0"/>
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

  <xsd:complexType name="CT_BubbleScale">
    <xsd:attribute name="val" type="ST_BubbleScale" default="100%"/>
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

  <xsd:complexType name="CT_SizeRepresents">
    <xsd:attribute name="val" type="ST_SizeRepresents" default="area"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Trendline">
    <xsd:sequence>
      <xsd:element name="name"          type="xsd:string"           minOccurs="0"/>
      <xsd:element name="spPr"          type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="trendlineType" type="CT_TrendlineType"/>
      <xsd:element name="order"         type="CT_Order"             minOccurs="0"/>
      <xsd:element name="period"        type="CT_Period"            minOccurs="0"/>
      <xsd:element name="forward"       type="CT_Double"            minOccurs="0"/>
      <xsd:element name="backward"      type="CT_Double"            minOccurs="0"/>
      <xsd:element name="intercept"     type="CT_Double"            minOccurs="0"/>
      <xsd:element name="dispRSqr"      type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="dispEq"        type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="trendlineLbl"  type="CT_TrendlineLbl"      minOccurs="0"/>
      <xsd:element name="extLst"        type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:simpleType name="ST_BubbleScale">
    <xsd:union memberTypes="ST_BubbleScalePercent ST_BubbleScaleUInt"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_BubbleScalePercent">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="0*(([0-9])|([1-9][0-9])|([1-2][0-9][0-9])|300)%"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_BubbleScaleUInt">
    <xsd:restriction base="xsd:unsignedInt">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="300"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_SizeRepresents">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="area"/>
      <xsd:enumeration value="w"/>
    </xsd:restriction>
  </xsd:simpleType>


References
----------

* https://blogs.msdn.microsoft.com/tomholl/2011/03/27/creating-multi-series-bubble-charts-in-excel/

* http://peltiertech.com/Excel/ChartsHowTo/HowToBubble.html

* http://peltiertech.com/Excel/Charts/ControlBubbleSizes.html
