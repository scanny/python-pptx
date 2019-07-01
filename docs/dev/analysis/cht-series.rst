
Chart - Series
==============

A series in a sequence of *points* that correspond to the categories of
a plot. A chart may have more than one series, which gives rise, for example,
to a *clustered column chart* or a line chart with multiple lines plotted.

A series belongs to a *plot*. When a chart has multiple plots, such as a bar
chart with a superimposed line plot, each of the series in the chart belong to
one plot or the other.

Although the Microsoft API has a method on Chart to access all the series in
a chart, it also has the same method for accessing the series of a plot
(ChartGroup in MS API).

There are eight distinct series types, corresponding to major chart types.
They all share a common base of attributes and others appear on one or more
of the types.


Series Access
-------------

Series are perhaps most naturally accessed from a plot, to which they
belong::

    >>> plot.series
    <pptx.chart.series.SeriesCollection instance at x11091e750>

However, there is also a property on |Chart| which allows access to all
the series in the chart::

    >>> chart.series
    <pptx.chart.series.SeriesCollection instance at x11091f970>

Each series in a chart has an explicit sequence indicator, the value of its
required `c:order` child element. The series for a plot appear in order of
this value. The series for a chart appear in plot order, then their order
within that plot, such that all series for the first plot appear before those
in the next plot, and so on.

Properties
----------

Series.format
~~~~~~~~~~~~~

All series have an optional `c:spPr` element that control the drawing shape
properties of the series such as fill and line, including transparency and
shadow.

Series.points
~~~~~~~~~~~~~

Perhaps counterintuitively, a Point object does not provide access to all the
attributes one might think. It only provides access to attributes of the
visual representation of the point in that chart, such as the color, datum
label, or marker. It does not provide access to the data point values, such
as the Y value or the bubble size.

Surface charts do not have a distinct data point representation (rather just
an inflection in the surface. So series of surface charts will not have the
.points member. Since surface charts are not yet implemented, this will come
into play sometime later.

Note that bubble and XY have a different way of organizing their data points
so have a distinct implementation from that of category charts.

**Implementation notes:**

* Introduce _BaseCategorySeries and subclass all category series types from
  it. Add tests to test inheritance. No acceptance test since this is
  internals-only.

* It's possible the only requirement is to create CategoryPoints. The rest of
  the implementation might work all on its own. Better spike it and see.


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
    >>> fill = series.format.fill
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

  <xsd:complexType name="CT_AreaSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"            type="CT_UnsignedInt"/>
      <xsd:element name="order"          type="CT_UnsignedInt"/>
      <xsd:element name="tx"             type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"           type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="pictureOptions" type="CT_PictureOptions"    minOccurs="0"/>
      <xsd:element name="dPt"            type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"          type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="trendline"      type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"        type="CT_ErrBars"           minOccurs="0" maxOccurs="2"/>
      <xsd:element name="cat"            type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"            type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
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

  <xsd:complexType name="CT_ScatterSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"       type="CT_UnsignedInt"/>
      <xsd:element name="order"     type="CT_UnsignedInt"/>
      <xsd:element name="tx"        type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"      type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="marker"    type="CT_Marker"            minOccurs="0"/>
      <xsd:element name="dPt"       type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"     type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="trendline" type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"   type="CT_ErrBars"           minOccurs="0" maxOccurs="2"/>
      <xsd:element name="xVal"      type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="yVal"      type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="smooth"    type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="extLst"    type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SurfaceSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"    type="CT_UnsignedInt"/>
      <xsd:element name="order"  type="CT_UnsignedInt"/>
      <xsd:element name="tx"     type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="cat"    type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"    type="CT_NumDataSource"     minOccurs="0"/>
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
