
Chart - Points
==============

A chart may be understood as a graph of a function, formed by rendering
a finite number of ordered pairs (e.g. (x, y)) in one of several graphical
formats such as bar, pie, or bubble.

A functon can be expressed as::

    y = f(x)

where *x* is the input to the function and *y* is the output. *x* is known
more formally as the function's *argument*, and *y* as its *value*. The set
of possible input (x) values is the *domain* of the function. The set of
possible output (y) values is the function's *range*.

In PowerPoint, the function is perhaps rarely mathematical; perhaps more
commonly, the x and y values are correlated observations, facts to be
communicated graphically as a set. Still these terms borrowed from
mathematics provide convenient language for describing aspects of PowerPoint
charts.

PowerPoint charts can be divided broadly into *category* types, which take
a discrete argument, and *XY* (aka. scatter) types which take a continuous
argument. For category types, the argument is one of a discrete set of
categories, expressed as a label, such as "Blue", "Red", and "Green". The
argument for an XY type is a real number. In both cases, the *value* of each
data point is a real number.

The argument is commonly associated with the horizontal axis of the chart and
the value with the vertical axis; these would be the X and Y axis
respectively on an XY chart. However the 2D orientation of the axes is
reversed in some char types (bar charts in particular) and takes other forms
in chart types such as radar and doughnut. In the MS API, the axis associated
with the argument is known as the *category axis* and the other is the *value
axis*. This terminology is used across all chart types (including XY) as
a matter of convention.


Series
------

A chart can have one or more series. A series is a set of ordered pairs that
are related in some way and meant to be distinguished graphically on the
chart, often so they can be compared. For example, one series could represent
Q1 financial results, depicted as a blue line, and a second series represent
Q2 financial results with a red line.

In a standard chart, all series share the same domain (e.g. the same X axis).
In a category chart, all series also share the same set of categories. In an
XY chart, the domain values are continuous; so in general, each data point in
each XY series will in have a distinct argument (x value).

In the MS API, there are multiple members on Series related to data points:

* Values
* XValues
* Points

Note that *Categories* is not in this set. ``CategoryCollection`` is a member
of ``ChartGroup`` (named Plot in |pp|) in the MS API.

**Values** is the Excel range address at which the series values are stored.
A set of constant values can be assigned in the MS API but this is not
supported in |pp|.

**XValues** is the Excel range address at which the series arguments are
stored when the chart is one of the XY types. X values replace categories for
an XY chart.

**Points** is a sequence of Point objects, each of which provides access to
the formatting of the graphical representation of the data point. Notably, it
does not allow changing the category, x or y value, or size (for a bubble
chart).

Every chart type supported by MS Office charts has the concept of data
points.

Informally, a data point is an atomic element of the content depicted by the
chart.

MS Office charts do not have a firm distinction for data points.

In all cases, points belong to a *series*.

There is some distinction between a data point for a category chart and one
for an XY (scatter) chart.

The API


Functions
---------

::

    data_point_count = min(
        xVal.ptCount_val, yVal.ptCount_val, bubbleSize.ptCount_val
    )

    invariant(ptCount_val > max(pt.idx))

    @property
    def pt_v(self, parent, idx):
        pts = parent.xpath('.//c:pt[@val=%s]' % idx)
        if pts is None:
            return None
        pt = pts[0]
        return pt.v

    points[i] ~= Point(xVal.pts[i], yVal.pts[i], bubbleSize.pts[i])


Point membership hypothesis:
----------------------------

* Visual alignment is not significant. A range may be anywhere in relation to
  the other data point ranges.

* Sequence is low row/col to high row/col in the Excel range, regardless of
  the "direction" in which the cells are selected. This seems to be enforced
  by a validation "correction" of the range string in the UI 'Select Data'
  dialog.

* The number (count) of x, y, or size values is the number of cells in the
  range. This corresponds directly to ptCount/@val for that data source
  sequence.

* The number of points in the series is the minimum of the x, y, and size
  value counts (ptCount).

* The ordered set of values for a point is formed by simple indexing within
  each value sequence. For example::

    point[3] = (x_values[3], y_values[3], sizes[3])

  When any value sequence in the set runs out of elements, no further points
  are formed.

  All values of each data source sequence are written to the XML; values are
  not truncated because they lack a counter part in one of the other
  sequences. Consequently, the ptCount values will not necessarily match.


Experiment (IronPython):
------------------------

* Create a chart with three X-values and four Y-values. .XValues has three
  and .Values has four members. How many points are in Series.Points?

Observations:

* 4 X, 4 Y, 4 size -- 4 points
* 3 X, 4 Y, 4 size -- 3 points
* 4 X, 3 Y, 4 size -- 3 points
* 4 X, 4 Y, 3 size -- 3 points

Points with blank (y) values still count as a point.
Points with blank (y) values still count as a point.

Explain how ...

Hypothesis: An x, y, or size value index always starts at zero, at the
beginning of the range, and increments to ptCount-1.

Hypothesis: ptCount is always based on the number of cells in the Excel
range, including blank cells at the start and end.

ptCount and pt behavior on range including blank cell at end.

xVal//ptCount can be different than yVal//ptCount


Related Schema Definitions
--------------------------

.. highlight:: xml

Point-related elements are stored under `c:ser`::

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

  <xsd:complexType name="CT_DLbls">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="dLbl"            type="CT_DLbl"              minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="numFmt"          type="CT_NumFmt"            minOccurs="0"/>
      <xsd:element name="spPr"            type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"            type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="dLblPos"         type="CT_DLblPos"           minOccurs="0"/>
      <xsd:element name="showLegendKey"   type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showVal"         type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showCatName"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showSerName"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showPercent"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showBubbleSize"  type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="separator"       type="xsd:string"           minOccurs="0"/>
      <xsd:element name="showLeaderLines" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="leaderLines"     type="CT_ChartLines"        minOccurs="0"/>
      <xsd:element name="extLst"          type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_DLbl">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"            type="CT_UnsignedInt"/>
      <xsd:element name="layout"         type="CT_Layout"            minOccurs="0"/>
      <xsd:element name="tx"             type="CT_Tx"                minOccurs="0"/>
      <xsd:element name="numFmt"         type="CT_NumFmt"            minOccurs="0"/>
      <xsd:element name="spPr"           type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"           type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="dLblPos"        type="CT_DLblPos"           minOccurs="0"/>
      <xsd:element name="showLegendKey"  type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showVal"        type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showCatName"    type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showSerName"    type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showPercent"    type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showBubbleSize" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="separator"      type="xsd:string"           minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
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
