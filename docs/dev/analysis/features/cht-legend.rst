
Chart - Legend
===================

A chart may have a legend. The legend may be placed within the plot area or
alongside it. The legend has a fill and font and may be positioned on the
top, right, bottom, left, in the upper-right corner, or in a custom position.
Individual legend entries may be custom specified, but that is not yet in
scope.


Candidate protocol
------------------

::

    >>> chart.has_legend
    False
    >>> chart.has_legend = True
    >>> chart.has_legend
    True
    >>> legend = chart.legend
    >>> legend
    <pptx.chart.chart.Legend object at 0xdeadbeef1>

    >>> legend.font
    <pptx.text.Font object at 0xdeadbeef2>

    >>> legend.horz_offset
    0.0
    >>> legend.horz_offset = 0.2
    >>> legend.horz_offset
    0.2

    >>> legend.include_in_layout
    True
    >>> legend.include_in_layout = False
    >>> legend.include_in_layout
    False

    >>> legend.position
    XL_LEGEND_POSITION.RIGHT (-4152)
    >>> legend.position = XL_LEGEND_POSITION.BOTTOM
    >>> legend.position
    XL_LEGEND_POSITION.BOTTOM (-4107)


Feature Summary
---------------

* :attr:`.Chart.has_legend` -- Read/write boolean property
* :attr:`.Chart.legend` -- Read-only |Legend| object or None
* :attr:`.Legend.horz_offset` -- Read/write float (-1.0 -> 1.0).
* :attr:`.Legend.include_in_layout` -- Read/write boolean.
* :attr:`.Legend.position` -- Read/write :ref:`XlLegendPosition`
* :attr:`.Legend.font` -- Read-only |Font| object.


Enumerations
------------

* :ref:`XlLegendPosition`


Microsoft API
-------------

Chart.HasLegend
    True if the chart has a legend. Read/write Boolean.

Chart.Legend
    Returns the legend for the chart. Read-only Legend.

Legend.IncludeInLayout
    True if a legend will occupy the chart layout space when a chart layout
    is being determined. The default is True. Read/write Boolean.

Legend.Position
    Returns or sets the position of the legend on the chart. Read/write
    XlLegendPosition.


XML specimens
-------------

.. highlight:: xml

Example legend XML::

  <c:legend>
    <c:legendPos val="t"/>
    <c:layout>
      <c:manualLayout>
        <c:xMode val="edge"/>
        <c:yMode val="edge"/>
        <c:x val="0.321245570866142"/>
        <c:y val="0.025"/>
        <c:w val="0.532508858267717"/>
        <c:h val="0.0854055118110236"/>
      </c:manualLayout>
    </c:layout>
    <c:overlay val="1"/>
    <c:spPr>
      <a:solidFill>
        <a:schemeClr val="accent6">
          <a:lumMod val="20000"/>
          <a:lumOff val="80000"/>
        </a:schemeClr>
      </a:solidFill>
    </c:spPr>
    <c:txPr>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
        <a:pPr>
          <a:defRPr sz="1600" b="0" i="1" baseline="0"/>
        </a:pPr>
        <a:endParaRPr lang="en-US"/>
      </a:p>
    </c:txPr>
  </c:legend>


Legend having horz_offset == 0.42::

  <c:legend>
    <c:legendPos val="r"/>
    <c:layout>
      <c:manualLayout>
        <c:xMode val="factor"/>
        <c:yMode val="factor"/>
        <c:x val="0.42"/>
      </c:manualLayout>
    </c:layout>
    <c:overlay val="0"/>
  </c:legend>


Related Schema Definitions
--------------------------

::

  <xsd:complexType name="CT_Legend">
    <xsd:sequence>
      <xsd:element name="legendPos"   type="CT_LegendPos"         minOccurs="0"/>
      <xsd:element name="legendEntry" type="CT_LegendEntry"       minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="layout"      type="CT_Layout"            minOccurs="0"/>
      <xsd:element name="overlay"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="spPr"        type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"        type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_LegendPos">
    <xsd:attribute name="val" type="ST_LegendPos" default="r"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_LegendPos">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="b"/>
      <xsd:enumeration value="tr"/>
      <xsd:enumeration value="l"/>
      <xsd:enumeration value="r"/>
      <xsd:enumeration value="t"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_LegendEntry">
    <xsd:sequence>
      <xsd:element name="idx" type="CT_UnsignedInt"/>
      <xsd:choice>
        <xsd:element name="delete" type="CT_Boolean"/>
        <xsd:group    ref="EG_LegendEntryData"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Layout">
    <xsd:sequence>
      <xsd:element name="manualLayout" type="CT_ManualLayout"  minOccurs="0"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_LayoutTarget">
    <xsd:attribute name="val" type="ST_LayoutTarget" default="outer"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_LayoutTarget">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="inner"/>
      <xsd:enumeration value="outer"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_LayoutMode">
    <xsd:attribute name="val" type="ST_LayoutMode" default="factor"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_LayoutMode">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="edge"/>
      <xsd:enumeration value="factor"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_ManualLayout">
    <xsd:sequence>
      <xsd:element name="layoutTarget" type="CT_LayoutTarget"  minOccurs="0"/>
      <xsd:element name="xMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="yMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="wMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="hMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="x"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="y"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="w"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="h"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>
