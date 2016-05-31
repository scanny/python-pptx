
Chart - InvertIfNegative
========================

Can be set on a bar series or on a data point. It's not clear what its meaning
is on a non-bar data point. ``c:barSer/c:invertIfNegative`` is the target
element. Also valid on a bubble series.


Protocol
--------

::

    >>> assert isinstance(plot, pptx.chart.plot.BarPlot)
    >>> series = plot.series[0]
    >>> series
    <pptx.chart.series.BarSeries instance at 0x1deadbeef>
    >>> series.invert_if_negative
    True
    >>> series.invert_if_negative = False
    >>> series.invert_if_negative
    False


Semantics
---------

* Defaults to |True| if the ``<c:invertIfNegative>`` element is not present.
* Defaults to |True| if the parent element is present but the `val` attribute
  is not.


PowerPoint® behavior
--------------------

In my tests, the ``<c:invertIfNegative>`` element is always present in a new
bar chart and always explicitly initialized to False.


XML specimens
-------------

.. highlight:: xml

A series element from a simple column chart (single series)::

  <c:ser>
    <c:idx val="0"/>
    <c:order val="0"/>
    <c:tx>
      <!-- ... -->
    </c:tx>
    <c:invertIfNegative val="0"/>
    <c:dLbls>
      <!-- ... -->
    </c:dLbls>
    <c:cat>
      <!-- ... -->
    </c:cat>
    <c:val>
      <!-- ... -->
    </c:val>
  </c:ser>


Related Schema Definitions
--------------------------

::

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
