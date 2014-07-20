
Chart - Series
==============

A series in a sequence of *points* that correspond to the categories of a plot.
A chart may have more than one series, which gives rise for example to
a *clustered column chart* or a line chart with multiple lines plotted.

A series belongs to a *plot*. When a chart has multiple plots, such as a bar
chart with a superimposed line plot, each of the series in the chart belong to
one plot or the other.

Although the Microsoft API has a method on Chart to access all the series in
a chart, it also has the same method for accessing the series of a plot
(ChartGroup in MS API).


Protocol
--------

::

    >>> assert isinstance(chart, pptx.chart.chart.Chart)
    >>> plot = chart.plots[0]
    >>> plot
    <pptx.chart.plot.BarChart instance at 0x1deadbeef>
    >>> series = plot.series[0]
    >>> series
    <pptx.chart.series.BarSeries instance at 0x...>
    >>> fill = series.fill
    >>> fill
    <pptx.dml.fill.FillFormat instance at 0x...>
    >>> fill.type
    None
    >>> fill.solid()
    >>> fill.fore_color.rgb = RGB(0x3F, 0x2C, 0x36)
    # also available are theme_color and brightness


XML specimens
-------------

.. highlight:: xml

A series element from a simple column chart (single series)::

  <c:ser>
    <c:idx val="0"/>
    <c:order val="0"/>
    <c:tx>
      <c:strRef>
        <c:f>Sheet1!$A$2</c:f>
        <c:strCache>
          <c:ptCount val="1"/>
          <c:pt idx="0">
            <c:v>Base</c:v>
          </c:pt>
        </c:strCache>
      </c:strRef>
    </c:tx>
    <c:invertIfNegative val="0"/>
    <c:dLbls>
      <c:numFmt formatCode="General" sourceLinked="0"/>
      <c:spPr>
        <a:noFill/>
        <a:ln>
          <a:noFill/>
        </a:ln>
        <a:effectLst/>
      </c:spPr>
      <c:txPr>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:pPr>
            <a:defRPr sz="1000"/>
          </a:pPr>
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </c:txPr>
      <c:dLblPos val="outEnd"/>
      <c:showLegendKey val="0"/>
      <c:showVal val="1"/>
      <c:showCatName val="0"/>
      <c:showSerName val="0"/>
      <c:showPercent val="0"/>
      <c:showBubbleSize val="0"/>
      <c:showLeaderLines val="0"/>
      <c:extLst xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"
                xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <c:ext xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"
               uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}">
          <c15:layout/>
          <c15:showLeaderLines val="0"/>
        </c:ext>
      </c:extLst>
    </c:dLbls>
    <c:cat>
      <c:strRef>
        <c:f>Sheet1!$B$1:$F$1</c:f>
        <c:strCache>
          <c:ptCount val="5"/>
          <c:pt idx="0">
            <c:v>Très probable</c:v>
          </c:pt>
          <c:pt idx="1">
            <c:v>Plutôt probable</c:v>
          </c:pt>
          <c:pt idx="2">
            <c:v>Plutôt improbable</c:v>
          </c:pt>
          <c:pt idx="3">
            <c:v>Très improbable</c:v>
          </c:pt>
          <c:pt idx="4">
            <c:v>Je ne sais pas</c:v>
          </c:pt>
        </c:strCache>
      </c:strRef>
    </c:cat>
    <c:val>
      <c:numRef>
        <c:f>Sheet1!$B$2:$F$2</c:f>
        <c:numCache>
          <c:formatCode>0</c:formatCode>
          <c:ptCount val="5"/>
          <c:pt idx="0">
            <c:v>19.0</c:v>
          </c:pt>
          <c:pt idx="1">
            <c:v>13.0</c:v>
          </c:pt>
          <c:pt idx="2">
            <c:v>10.0</c:v>
          </c:pt>
          <c:pt idx="3">
            <c:v>46.0</c:v>
          </c:pt>
          <c:pt idx="4">
            <c:v>12.0</c:v>
          </c:pt>
        </c:numCache>
      </c:numRef>
    </c:val>
  </c:ser>


Related Schema Definitions
--------------------------

::

  <xsd:group name="EG_SerShared">
    <xsd:sequence>
      <xsd:element name="idx"   type="CT_UnsignedInt"/>
      <xsd:element name="order" type="CT_UnsignedInt"/>
      <xsd:element name="tx"    type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"  type="a:CT_ShapeProperties" minOccurs="0"/>
    </xsd:sequence>
  </xsd:group>

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

  <xsd:complexType name="CT_BarSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"              type="CT_UnsignedInt"/>
      <xsd:element name="order"            type="CT_UnsignedInt"/>
      <xsd:element name="tx"               type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"             type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="invertIfNegative" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="pictureOptions"   type="CT_PictureOptions"    minOccurs="0"/>
      <xsd:element name="dPt"              type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"            type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="trendline"        type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"          type="CT_ErrBars"           minOccurs="0"/>
      <xsd:element name="cat"              type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"              type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="shape"            type="CT_Shape"             minOccurs="0"/>
      <xsd:element name="extLst"           type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ScatterSer">
    <xsd:sequence>
      <xsd:group   ref="EG_SerShared"/>
      <xsd:element name="marker"    type="CT_Marker"        minOccurs="0"/>
      <xsd:element name="dPt"       type="CT_DPt"           minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"     type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="trendline" type="CT_Trendline"     minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"   type="CT_ErrBars"       minOccurs="0" maxOccurs="2"/>
      <xsd:element name="xVal"      type="CT_AxDataSource"  minOccurs="0"/>
      <xsd:element name="yVal"      type="CT_NumDataSource" minOccurs="0"/>
      <xsd:element name="smooth"    type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="extLst"    type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_RadarSer">
    <xsd:sequence>
      <xsd:group   ref="EG_SerShared"/>
      <xsd:element name="marker" type="CT_Marker"        minOccurs="0"/>
      <xsd:element name="dPt"    type="CT_DPt"           minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"  type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="cat"    type="CT_AxDataSource"  minOccurs="0"/>
      <xsd:element name="val"    type="CT_NumDataSource" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_BarSer">
    <xsd:sequence>
      <xsd:group   ref="EG_SerShared"/>
      <xsd:element name="invertIfNegative" type="CT_Boolean"        minOccurs="0"/>
      <xsd:element name="pictureOptions"   type="CT_PictureOptions" minOccurs="0"/>
      <xsd:element name="dPt"              type="CT_DPt"            minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"            type="CT_DLbls"          minOccurs="0"/>
      <xsd:element name="trendline"        type="CT_Trendline"      minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"          type="CT_ErrBars"        minOccurs="0"/>
      <xsd:element name="cat"              type="CT_AxDataSource"   minOccurs="0"/>
      <xsd:element name="val"              type="CT_NumDataSource"  minOccurs="0"/>
      <xsd:element name="shape"            type="CT_Shape"          minOccurs="0"/>
      <xsd:element name="extLst"           type="CT_ExtensionList"  minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_AreaSer">
    <xsd:sequence>
      <xsd:group ref="EG_SerShared"/>
      <xsd:element name="pictureOptions" type="CT_PictureOptions" minOccurs="0"/>
      <xsd:element name="dPt"            type="CT_DPt"            minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"          type="CT_DLbls"          minOccurs="0"/>
      <xsd:element name="trendline"      type="CT_Trendline"      minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"        type="CT_ErrBars"        minOccurs="0" maxOccurs="2"/>
      <xsd:element name="cat"            type="CT_AxDataSource"   minOccurs="0"/>
      <xsd:element name="val"            type="CT_NumDataSource"  minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"  minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_PieSer">
    <xsd:sequence>
      <xsd:group ref="EG_SerShared"/>
      <xsd:element name="explosion" type="CT_UnsignedInt"   minOccurs="0"/>
      <xsd:element name="dPt"       type="CT_DPt"           minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"     type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="cat"       type="CT_AxDataSource"  minOccurs="0"/>
      <xsd:element name="val"       type="CT_NumDataSource" minOccurs="0"/>
      <xsd:element name="extLst"    type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_BubbleSer">
    <xsd:sequence>
      <xsd:group ref="EG_SerShared"/>
      <xsd:element name="invertIfNegative" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="dPt"              type="CT_DPt"           minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"            type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="trendline"        type="CT_Trendline"     minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"          type="CT_ErrBars"       minOccurs="0" maxOccurs="2"/>
      <xsd:element name="xVal"             type="CT_AxDataSource"  minOccurs="0"/>
      <xsd:element name="yVal"             type="CT_NumDataSource" minOccurs="0"/>
      <xsd:element name="bubbleSize"       type="CT_NumDataSource" minOccurs="0"/>
      <xsd:element name="bubble3D"         type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="extLst"           type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SurfaceSer">
    <xsd:sequence>
      <xsd:group ref="EG_SerShared"/>
      <xsd:element name="cat"    type="CT_AxDataSource"  minOccurs="0"/>
      <xsd:element name="val"    type="CT_NumDataSource" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>
