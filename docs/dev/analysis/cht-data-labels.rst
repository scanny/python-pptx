
Chart - Data Labels
===================

On a PowerPoint chart, data points may be labeled as an aid to readers.
Typically, the label is the value of the data point, but a data label may
have any combination of its series name, category name, and value. A number
format may also be applied to the value displayed.


Object access
-------------

The `DataLabels` object is not a collection of `DataLabel` objects.
`DataLabels` controls the formatting of the data labels for a whole series,
so "global" settings in a way of thinking. An individual `DataLabel` object
is accessed using it's associated `Point` object::

    data_label = chart.plot[0].series[0].points[0].data_label

The two object types share many attributes.


DataLabel.font
--------------

Proposed protocol::

    >>> font = data_label.font
    >>> font.color.rgb = RGBColor(0x3C, 0xEB, 0x27)


Position
--------

There are nine choices for where a data label may be positioned relative to
its data point marker, although the available options depend on the chart
type. The options are specified using the :ref:`XlDataLabelPosition`
enumeration.

**XML Semantics.** The default position when no ``<c:dLblPos>`` element is
present (common) depends on the chart type:

+------------------------------+-------------+
| barChart (clustered)         | OUTSIDE_END |
+------------------------------+-------------+
| bar3DChart (clustered)       | OUTSIDE_END |
+------------------------------+-------------+
| barChart (stacked)           | CENTER      |
+------------------------------+-------------+
| barChart (percent stacked)   | CENTER      |
+------------------------------+-------------+
| bar3DChart (stacked)         | CENTER      |
+------------------------------+-------------+
| bar3DChart (percent stacked) | CENTER      |
+------------------------------+-------------+
| pieChart                     | BEST_FIT    |
+------------------------------+-------------+
| pie3DChart                   | BEST_FIT    |
+------------------------------+-------------+
| ofPieChart                   | BEST_FIT    |
+------------------------------+-------------+
| areaChart                    | CENTER      |
+------------------------------+-------------+
| area3DChart                  | CENTER      |
+------------------------------+-------------+
| doughnutChart                | CENTER      |
+------------------------------+-------------+
| radarChart                   | OUTSIDE_END |
+------------------------------+-------------+
| all others                   | RIGHT       |
+------------------------------+-------------+

http://msdn.microsoft.com/en-us/library/ff535061(v=office.12).aspx

Proposed protocol::

    >>> data_labels = plot.data_labels
    >>> data_labels.position
    OUTSIDE_END (2)
    >>> data_labels.position = XL_DATA_LABEL_POSITION.INSIDE_END
    >>> data_labels.position
    INSIDE_END (3)


MS API
------

The following properties from the MS API DataLabel object are of most
interest:

Delete()
    Deletes the DataLabel object, which might be useful for "undoing" data
    label overrides and restoring inherited behaviors.

AutoText
    True if the object automatically generates appropriate text based on
    context. Read/write Boolean.

Caption
    Returns or sets the data label text. Read/write String.

Format
    Returns the line, fill, and effect formatting for the object. Read-only
    ChartFormat.

NumberFormat
    Returns or sets the format code for the object. Read/write String.

NumberFormatLinked
    True if the number format is linked to the cells (so that the number
    format changes in the labels when it changes in the cells). Read/write
    Boolean.

Orientation
    Returns or sets the text orientation. Read/write Long in range -90 to 90
    (degrees).

Position
    Returns or sets the position of the data label. Read/write
    XlDataLabelPosition.

Separator
    Returns or sets the separator used for the data labels on a chart. Used
    when more than one value is shown, like category+value. Read/write
    Variant.

Shadow
    Returns or sets a value that indicates whether the object has a shadow.
    Read/write Boolean.

ShowBubbleSize
    True to show the bubble size for the data labels on a chart. False to
    hide the bubble size. Read/write Boolean.

ShowCategoryName
    True to display the category name for the data labels on a chart. False
    to hide the category name. Read/write Boolean.

ShowLegendKey
    True if the data label legend key is visible. This is a small square of
    the series color shown in the legend, and appears adjacent to the data
    label. Read/write Boolean.

ShowPercentage
    True to display the percentage value for the data labels on a chart.
    False to hide the value. This might only be relevant for pie charts and
    similar. Read/write Boolean.

ShowSeriesName
    True to show the series name for the data labels on a chart. False to
    hide the series name. Read/write Boolean.

ShowValue
    True to display a specified chart's data label values. False to hide the
    values. Read/write Boolean.

Text
    Returns or sets the text for the specified object. Read/write String.

VerticalAlignment
    Returns or sets the vertical alignment of the specified object.
    Read/write Variant.


PowerPoint behavior
-------------------

* A default PowerPoint bar chart does not display data labels, but it does
  have a ``<c:dLbls>`` child element on its ``<c:barChart>`` element.

* Data labels are added to a chart in the UI by selecting the *Data Labels*
  drop-down menu in the Chart Layout ribbon. The options include setting the
  contents of the data label, its position relative to the point, and
  bringing up the *Format Data Labels* dialog.

* The default number format, when no ``<c:numFmt>`` child element appears, is
  equivalent to ``<c:numFmt formatCode="General" sourceLinked="1"/>``


XML Semantics
-------------

* A ``<c:dLbls>`` element at the plot level (e.g. ``<c:barChart>``) is
  overridden completely by a ``<c:dLbls>`` element at the series level.
  Unless overridden, a ``<c:dLbls>`` element at the plot level determines the
  content and formatting for data labels on all the plot's series.

* A ``<c:dLbl>`` element appears as a child of ``<c:dLbls>`` only when the
  data label for its associated data point has overrides.


XML specimens
-------------

.. highlight:: xml

The ``<c:dLbls>`` element is available on a plot (e.g. ``<c:barChart>``),
a series (``<c:ser>``), and perhaps elsewhere.

Default ``<c:dLbls>`` element added by PowerPoint when selecting Data Labels
> Value from the Chart Layout ribbon::

    <c:dLbls>
      <c:showLegendKey val="0"/>
      <c:showVal val="1"/>
      <c:showCatName val="0"/>
      <c:showSerName val="0"/>
      <c:showPercent val="0"/>
      <c:showBubbleSize val="0"/>
    </c:dLbls>

A ``<c:dLbls>`` element specifying the labels should appear in 10pt Bold
Italic Arial Narrow, color Accent 6, 25% Darker::

    <c:dLbls>
      <c:txPr>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:pPr>
            <a:defRPr sz="1000" b="1" i="1">
              <a:solidFill>
                <a:schemeClr val="accent6">
                  <a:lumMod val="75000"/>
                </a:schemeClr>
              </a:solidFill>
              <a:latin typeface="Arial Narrow"/>
            </a:defRPr>
          </a:pPr>
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </c:txPr>
      <c:showLegendKey val="0"/>
      <c:showVal val="1"/>
      <c:showCatName val="0"/>
      <c:showSerName val="0"/>
      <c:showPercent val="0"/>
      <c:showBubbleSize val="0"/>
    </c:dLbls>

A ``<c:dLbls>`` element having an individual point override for font color
and outline box::

    <c:dLbls>
      <c:dLbl>
        <c:idx val="0"/>
        <c:spPr>
          <!-- ouline color (Red) -->
          <a:ln>
            <a:solidFill>
              <a:srgbClr val="FF0000"/>
            </a:solidFill>
          </a:ln>
        </c:spPr>
        <c:txPr>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <!-- font color (Green) -->
              <a:defRPr>
                <a:solidFill>
                  <a:srgbClr val="00FF00"/>
                </a:solidFill>
              </a:defRPr>
            </a:pPr>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </c:txPr>
        <c:showLegendKey val="0"/>
        <c:showCatName val="0"/>
        <c:showSerName val="0"/>
      </c:dLbl>
      <c:spPr/>
      <!-- ... -->
    </c:dLbls>


Related Schema Definitions
--------------------------

::

  <xsd:complexType name="CT_DLbls">
    <xsd:sequence>
      <xsd:element name="dLbl" type="CT_DLbl" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:choice>
        <xsd:element name="delete"      type="CT_Boolean"/>
        <xsd:group   ref="Group_DLbls"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="Group_DLbls">  <!-- denormalized -->
    <xsd:sequence>
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
    </xsd:sequence>
  </xsd:group>

  <xsd:complexType name="CT_DLbl">
    <xsd:sequence>
      <xsd:element name="idx" type="CT_UnsignedInt"/>
      <xsd:choice>
        <xsd:element name="delete"     type="CT_Boolean"/>
        <xsd:group   ref="Group_DLbl"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="Group_DLbl"> <!-- denormalized -->
    <xsd:sequence>
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
    </xsd:sequence>
  </xsd:group>

  <xsd:complexType name="CT_DLblPos">
    <xsd:attribute name="val" type="ST_DLblPos" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_NumFmt">
    <xsd:attribute name="formatCode"   type="xsd:string"  use="required"/>
    <xsd:attribute name="sourceLinked" type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TextBody">
    <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle" minOccurs="0"/>
      <xsd:element name="p"        type="CT_TextParagraph" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Tx">
    <xsd:sequence>
      <xsd:choice minOccurs="1" maxOccurs="1">
        <xsd:element name="strRef" type="CT_StrRef"/>
        <xsd:element name="rich"   type="a:CT_TextBody"/>
      </xsd:choice>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:simpleType name="ST_DLblPos">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="bestFit"/>
      <xsd:enumeration value="b"/>
      <xsd:enumeration value="ctr"/>
      <xsd:enumeration value="inBase"/>
      <xsd:enumeration value="inEnd"/>
      <xsd:enumeration value="l"/>
      <xsd:enumeration value="outEnd"/>
      <xsd:enumeration value="r"/>
      <xsd:enumeration value="t"/>
    </xsd:restriction>
  </xsd:simpleType>
