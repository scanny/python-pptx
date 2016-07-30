
Chart - Markers
===============

A marker is a small geometric shape that explicitly indicates a data point
position on a line-oriented chart. Line, XY, and Radar are the line-oriented
charts. Other chart types do not have markers.


PowerPoint behavior
-------------------

At series level, UI provides property controls for:

* Marker Fill

  + Solid
  + Gradient
  + Picture or Texture
  + Pattern

* Marker Line

  + Solid

    - RGBColor (probably HSB too, etc.)
    - Transparency

  + Gradient
  + Weights & Arrows

* Marker Style

  + Choice of enumerated shapes
  + Size


XL_MARKER_STYLE enumeration
---------------------------

https://msdn.microsoft.com/en-us/library/bb241374(v=office.12).aspx

+------------------------+-------+----------------------------------+
| xlMarkerStyleAutomatic | -4105 | Automatic markers                |
+------------------------+-------+----------------------------------+
| xlMarkerStyleCircle    | 8     | Circular markers                 |
+------------------------+-------+----------------------------------+
| xlMarkerStyleDash      | -4115 | Long bar markers                 |
+------------------------+-------+----------------------------------+
| xlMarkerStyleDiamond   | 2     | Diamond-shaped markers           |
+------------------------+-------+----------------------------------+
| xlMarkerStyleDot       | -4118 | Short bar markers                |
+------------------------+-------+----------------------------------+
| xlMarkerStyleNone      | -4142 | No markers                       |
+------------------------+-------+----------------------------------+
| xlMarkerStylePicture   | -4147 | Picture markers                  |
+------------------------+-------+----------------------------------+
| xlMarkerStylePlus      | 9     | Square markers with a plus sign  |
+------------------------+-------+----------------------------------+
| xlMarkerStyleSquare    | 1     | Square markers                   |
+------------------------+-------+----------------------------------+
| xlMarkerStyleStar      | 5     | Square markers with an  asterisk |
+------------------------+-------+----------------------------------+
| xlMarkerStyleTriangle  | 3     | Triangular markers               |
+------------------------+-------+----------------------------------+
| xlMarkerStyleX         | -4168 | Square markers with an X         |
+------------------------+-------+----------------------------------+

XML specimens
-------------

.. highlight:: xml

Marker properties set at the series level, all markers for the series::

  <c:ser>
    <!-- ... -->
    <c:marker>
      <c:spPr>
        <a:solidFill>
          <a:schemeClr val="accent4">
            <a:lumMod val="60000"/>
            <a:lumOff val="40000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:ln>
          <a:solidFill>
            <a:schemeClr val="accent6">
              <a:alpha val="51000"/>
            </a:schemeClr>
          </a:solidFill>
        </a:ln>
      </c:spPr>
    </c:marker>
    <!-- ... -->
  </c:ser>

Marker properties set on an individual point::

  <c:ser>
    <!-- ... -->
    <c:dPt>
      <c:idx val="0"/>
      <c:marker>
        <c:symbol val="circle"/>
        <c:size val="16"/>
        <c:spPr>
          <a:gradFill flip="none" rotWithShape="1">
            <a:gsLst>
              <a:gs pos="0">
                <a:schemeClr val="accent2"/>
              </a:gs>
              <a:gs pos="80000">
                <a:srgbClr val="FFFFFF">
                  <a:alpha val="61000"/>
                </a:srgbClr>
              </a:gs>
            </a:gsLst>
            <a:path path="circle">
              <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
            </a:path>
            <a:tileRect/>
          </a:gradFill>
          <a:ln w="12700">
            <a:solidFill>
              <a:schemeClr val="accent6"/>
            </a:solidFill>
            <a:prstDash val="sysDot"/>
          </a:ln>
          <a:effectLst>
            <a:outerShdw blurRad="50800" dist="38100" dir="14160000" algn="tl" rotWithShape="0">
              <a:schemeClr val="accent6">
                <a:lumMod val="75000"/>
                <a:alpha val="43000"/>
              </a:schemeClr>
            </a:outerShdw>
          </a:effectLst>
        </c:spPr>
      </c:marker>
      <c:bubble3D val="0"/>
      <c:spPr>
        <a:effectLst>
          <a:outerShdw blurRad="50800" dist="38100" dir="14160000" algn="tl" rotWithShape="0">
            <a:schemeClr val="accent6">
              <a:lumMod val="75000"/>
              <a:alpha val="43000"/>
            </a:schemeClr>
          </a:outerShdw>
        </a:effectLst>
      </c:spPr>
    </c:dPt>
    <!-- ... -->
  </c:ser>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

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
