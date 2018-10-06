
Line format
===========


Initial scope
-------------

* Shape.line, Connector.line to follow, then Cell.border(s).line, possibly
  text underline and font.line (character outline)


Protocol
--------

**Accessing line.** A shape object has a `.line` property that returns
a `LineFormat` object. The `.line` property is idempotent; it always returns
the same `LineFormat` object for a given shape object::

    >>> line = shape.line
    >>> assert(isinstance(line, LineFormat)

**Line properties.** A line has multiple properties, such as width, dash style, head and tail ends and fill::

    >>> shape.line
    <pptx.dml.line.LineFormat at 0x1b8f6f86eb8>
    >>> shape.line.color
    <pptx.dml.color.ColorFormat at 0x1b8f6f83908>
    >>> shape.line.width = Pt(1)
    >>> shape.line.width
    12700

Dash style can be set using `MSO_LINE_DASH_STYLE`::

    >>> shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    >>> shape.line.dash_style
    4

Tail and head ends style can be set using `MSO_ARROWHEAD_STYLE`::

    >>> shape.line.tail_end = MSO_ARROWHEAD_STYLE.DIAMOND
    >>> shape.line.tail_end
    5
    >>> shape.line.tail_end = MSO_ARROWHEAD_STYLE.OVAL
    >>> shape.line.tail_end
    6


Notes
-----

definitely spike on this one.

start with line.color, first an object to hang members off of


Issue: How to accommodate these competing requirements:

* Polymorphism in parent shape. More than one type of shape can have a line
  and possibly several (table cell border, Run/text, and text underline use
  the same CT_LineProperties element, and perhaps others).

* The line object cannot hold onto a `<a:ln>` element, even if that was
  a good idea, because it is an optional child; not having an `<a:ln>`
  element is a legitimate and common situation, indicating line formatting
  should be inherited from the theme or perhaps a layout placeholder.

* Needing to accommodate XML changing might not be important, could make that
  operation immutable, such that changing the shape XML returns a new shape,
  not changing the existing one in-place.

* maybe having the following 'line_format_owner_interface', delegating
  create, read, and delete of the `<a:ln>` element to the parent, and
  allowing LineFormat to take responsibility for update.

  + line.parent has the shape having the line format
  + parent.ln has the `<a:ln>` element or None, delegating access to the parent
  + ln = parent.add_ln()
  + parent.remove_ln()


MS API
------

* | LineFormat
  | https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/lineformat-object-powerpoint

* | LineFormat.DashStyle
  | https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/lineformat-dashstyle-property-powerpoint

* | MsoLineDashStyle Enumeration
  | https://docs.microsoft.com/en-us/office/vba/api/office.msolinedashstyle

  +-----------------------+-------+----------------------------------+
  | Name                  | Value | Description                      |
  +-----------------------+-------+----------------------------------+
  | msoLineDash           | 4     | Line consists of dashes only.    |
  +-----------------------+-------+----------------------------------+
  | msoLineDashDot        | 5     | Line is a dash-dot pattern.      |
  +-----------------------+-------+----------------------------------+
  | msoLineDashDotDot     | 6     | Line is a dash-dot-dot pattern.  |
  +-----------------------+-------+----------------------------------+
  | msoLineDashStyleMixed | -2    | Not supported.                   |
  +-----------------------+-------+----------------------------------+
  | msoLineLongDash       | 7     | Line consists of long dashes.    |
  +-----------------------+-------+----------------------------------+
  | msoLineLongDashDot    | 8     | Line is a long dash-dot pattern. |
  +-----------------------+-------+----------------------------------+
  | msoLineRoundDot       | 3     | Line is made up of round dots.   |
  +-----------------------+-------+----------------------------------+
  | msoLineSolid          | 1     | Line is solid.                   |
  +-----------------------+-------+----------------------------------+
  | msoLineSquareDot      | 2     | Line is made up of square dots.  |
  +-----------------------+-------+----------------------------------+

* | MsoArrowheadStyle Enumeration
  | https://docs.microsoft.com/en-us/office/vba/api/Office.MsoArrowheadStyle

  +------------------------+-------+---------------------------------+
  | Name                   | Value | Description                     |
  +------------------------+-------+---------------------------------+
  | msoArrowheadDiamond    | 5     | Diamond-shaped.                 |
  +------------------------+-------+---------------------------------+
  | msoArrowheadNone       | 1     | No arrowhead.                   |
  +------------------------+-------+---------------------------------+
  | msoArrowheadOpen       | 3     | Open.                           |
  +------------------------+-------+---------------------------------+
  | msoArrowheadOval       | 6     | Oval-shaped.                    |
  +------------------------+-------+---------------------------------+
  | msoArrowheadStealth    | 4     | Stealth-shaped.                 |
  +------------------------+-------+---------------------------------+
  | msoArrowheadStyleMixed | -2    | Return value only.              |
  +------------------------+-------+---------------------------------+
  | msoArrowheadTriangle   | 5     | Triangular.                     |
  +------------------------+-------+---------------------------------+

* | MsoArrowheadLength Enumeration
  | https://docs.microsoft.com/en-us/office/vba/api/Office.MsoArrowheadLength

  +--------------------------+-------+---------------------------------+
  | Name                     | Value | Description                     |
  +--------------------------+-------+---------------------------------+
  | msoArrowheadLengthMedium | 2     | Medium.                         |
  +--------------------------+-------+---------------------------------+
  | msoArrowheadLengthMixed  | -2    | Return value only.              |
  +--------------------------+-------+---------------------------------+
  | msoArrowheadLong         | 3     | Long.                           |
  +--------------------------+-------+---------------------------------+
  | msoArrowheadShort        | 1     | Short.                          |
  +--------------------------+-------+---------------------------------+

* | MsoArrowheadWidth Enumeration
  | https://docs.microsoft.com/en-us/office/vba/api/Office.MsoArrowheadWidth

  +-------------------------+-------+---------------------------------+
  | Name                    | Value | Description                     |
  +-------------------------+-------+---------------------------------+
  | msoArrowheadNarrow      | 1     | Narrow.                         |
  +-------------------------+-------+---------------------------------+
  | msoArrowheadWide        | 3     | Wide.                           |
  +-------------------------+-------+---------------------------------+
  | msoArrowheadWidthMedium | 2     | Medium.                         |
  +-------------------------+-------+---------------------------------+
  | msoArrowheadWidthMixed  | -2    | Return value only.              |
  +-------------------------+-------+---------------------------------+


Specimen XML
------------

.. highlight:: xml

Default 1 Pt line::

    <a:ln w="12700"/>

Solid line color::

    <p:spPr>
      <a:xfrm>
        <a:off x="950964" y="2925277"/>
        <a:ext cx="1257921" cy="619967"/>
      </a:xfrm>
      <a:prstGeom prst="curvedConnector3">
        <a:avLst/>
      </a:prstGeom>
      <a:ln>
        <a:solidFill>
          <a:schemeClr val="accent2"/>
        </a:solidFill>
      </a:ln>
    </p:spPr>

Line ends::

    <a:ln w="12700">
      <a:headEnd type="oval" w="med" len="lg"/>
      <a:tailEnd type="diamond" w="sm" len="med"/>
    </a:ln>

Little bit of everything::

    <p:spPr>
      <a:xfrm>
        <a:off x="950964" y="1101493"/>
        <a:ext cx="1257921" cy="0"/>
      </a:xfrm>
      <a:prstGeom prst="line">
        <a:avLst/>
      </a:prstGeom>
      <a:ln w="57150" cap="rnd" cmpd="thickThin">
        <a:gradFill flip="none" rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="accent1"/>
            </a:gs>
            <a:gs pos="100000">
              <a:srgbClr val="FFFFFF"/>
            </a:gs>
          </a:gsLst>
          <a:lin ang="0" scaled="1"/>
          <a:tileRect/>
        </a:gradFill>
        <a:prstDash val="sysDash"/>
        <a:bevel/>
        <a:headEnd type="oval"/>
        <a:tailEnd type="diamond"/>
      </a:ln>
    </p:spPr>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"    type="CT_Transform2D"            minOccurs="0"/>
      <xsd:group   ref="EG_Geometry"                               minOccurs="0"/>
      <xsd:group   ref="EG_FillProperties"                         minOccurs="0"/>
      <xsd:element name="ln"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:group   ref="EG_EffectProperties"                       minOccurs="0"/>
      <xsd:element name="scene3d" type="CT_Scene3D"                minOccurs="0"/>
      <xsd:element name="sp3d"    type="CT_Shape3D"                minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_LineProperties">
    <xsd:sequence>
      <xsd:group   ref="EG_LineFillProperties"                     minOccurs="0"/>
      <xsd:group   ref="EG_LineDashProperties"                     minOccurs="0"/>
      <xsd:group   ref="EG_LineJoinProperties"                     minOccurs="0"/>
      <xsd:element name="headEnd" type="CT_LineEndProperties"      minOccurs="0"/>
      <xsd:element name="tailEnd" type="CT_LineEndProperties"      minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="w"    type="ST_LineWidth"/>
    <xsd:attribute name="cap"  type="ST_LineCap"/>
    <xsd:attribute name="cmpd" type="ST_CompoundLine"/>
    <xsd:attribute name="algn" type="ST_PenAlignment"/>
  </xsd:complexType>

  <xsd:complexType name="CT_LineEndProperties">
    <xsd:attribute name="type" type="ST_LineEndType" use="optional"/>
    <xsd:attribute name="w" type="ST_LineEndWidth" use="optional"/>
    <xsd:attribute name="len" type="ST_LineEndLength" use="optional"/>
  </xsd:complexType>

  <xsd:group name="EG_LineFillProperties">
    <xsd:choice>
      <xsd:element name="noFill"    type="CT_NoFillProperties"/>
      <xsd:element name="solidFill" type="CT_SolidColorFillProperties"/>
      <xsd:element name="gradFill"  type="CT_GradientFillProperties"/>
      <xsd:element name="pattFill"  type="CT_PatternFillProperties"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_LineDashProperties">
    <xsd:choice>
      <xsd:element name="prstDash" type="CT_PresetLineDashProperties"/>
      <xsd:element name="custDash" type="CT_DashStopList"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_PresetLineDashProperties">
    <xsd:attribute name="val" type="ST_PresetLineDashVal" use="optional"/>
  </xsd:complexType>

  <xsd:group name="EG_LineJoinProperties">
    <xsd:choice>
      <xsd:element name="round" type="CT_LineJoinRound"/>
      <xsd:element name="bevel" type="CT_LineJoinBevel"/>
      <xsd:element name="miter" type="CT_LineJoinMiterProperties"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_EffectProperties">
    <xsd:choice>
      <xsd:element name="effectLst" type="CT_EffectList"/>
      <xsd:element name="effectDag" type="CT_EffectContainer"/>
    </xsd:choice>
  </xsd:group>

  <xsd:simpleType name="ST_LineEndType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="triangle"/>
      <xsd:enumeration value="stealth"/>
      <xsd:enumeration value="diamond"/>
      <xsd:enumeration value="oval"/>
      <xsd:enumeration value="arrow"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_LineEndWidth">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="sm"/>
      <xsd:enumeration value="med"/>
      <xsd:enumeration value="lg"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_LineEndLength">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="sm"/>
      <xsd:enumeration value="med"/>
      <xsd:enumeration value="lg"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_LineWidth">
    <xsd:restriction base="ST_Coordinate32Unqualified">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="20116800"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Coordinate32Unqualified">
    <xsd:restriction base="xsd:int"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_PresetLineDashVal">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="solid"/>
      <xsd:enumeration value="dot"/>
      <xsd:enumeration value="dash"/>
      <xsd:enumeration value="lgDash"/>
      <xsd:enumeration value="dashDot"/>
      <xsd:enumeration value="lgDashDot"/>
      <xsd:enumeration value="lgDashDotDot"/>
      <xsd:enumeration value="sysDash"/>
      <xsd:enumeration value="sysDot"/>
      <xsd:enumeration value="sysDashDot"/>
      <xsd:enumeration value="sysDashDotDot"/>
    </xsd:restriction>
  </xsd:simpleType>
