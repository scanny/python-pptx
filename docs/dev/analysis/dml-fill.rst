
Fill (for shapes)
=================

Overview
--------

In simple terms, *fill* is the color of a shape. The term *fill* is used to
distinguish from *line*, the border around a shape, which can have a different
color. It turns out that a line has a fill all its own, but that's another
topic.

While having a solid color fill is perhaps most common, it is just one of
several possible fill types. A shape can have the following types of fill:

* solid (color)
* gradient -- a smooth transition between multiple colors
* picture -- in which an image "shows through" and is cropped by the boundaries
  of the shape
* pattern -- in which a small square of pixels (tile) is repeated to fill the
  shape
* background (no fill) -- the body of the shape is transparent, allowing
  whatever is behind it to show through.

Elements that can have a fill include autoshape, line, text (treating each
glyph as a shape), table, table cell, and slide background.


Scope
-----

This analysis focuses on solid fill for an autoshape (including text box).
This relatively narrow initial focus is to enable getting to feature
implementation as quickly as possible, while understanding enough to design the
top-level of the API in a way that will allow other fill types to be added
incrementally.


Minimum viable feature
----------------------

Start with solid fill since that's the most frequently asked for and it reuses
the work completed on ColorFormat for font color feature::

    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0x01, 0x23, 0x45)
    fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    fill.fore_color.brightness = 0.25
    fill.transparency = 0.25
    shape.fill = None
    fill.background()  # is almost free once the rest is in place


Protocol
--------

**Accessing fill.** A shape object unconditionally has a `FillFormat` object
on `.fill`. The `.fill` property is idempotent; it always returns the same
`FillFormat` object for a given shape object::

    >>> fill = shape.fill
    >>> assert(isinstance(fill, FillFormat)

**Fill type.** A fill has a type, which may be |None|. The fill type
partially determines the valid calls on the fill::

    >>> fill.type
    None
    >>> fill.solid()
    >>> fill.type
    1  # MSO_FILL.SOLID

    >>> fill.fore_color = 'anything'
    AttributeError: can't set attribute  # .fore_color is read-only
    >>> fore_color = fill.fore_color
    >>> assert(isinstance(fore_color, ColorFormat))

    >>> fore_color.rgb = RGBColor(0x3F, 0x2c, 0x36)
    >>> fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    >>> fore_color.brightness = -0.25

    >>> fill.transparency
    0.0
    >>> fill.transparency = 0.25  # sets opacity to 75%

    >>> sp.fill = None  # removes any fill, fill is inherited from theme

**Pattern Fill.** ::

    >>> fill = shape.fill
    >>> fill.type
    None
    >>> fill.patterned()
    >>> fill.type
    2  # MSO_FILL.PATTERNED
    >>> fill.fore_color.rgb = RGBColor(79, 129, 189)
    >>> fill.back_color.rgb = RGBColor(239, 169, 6)


*fill type get/set*

* ``FillFormat.type`` -- ``MSO_FILL.SOLID`` or |None| for a start
* ``FillFormat.solid()`` -- changes fill type to <a:solidFill>
* ``FillFormat.fore_color`` -- changes fill type to <a:solidFill>
* ``FillFormat.transparency`` -- adds/changes <a:alpha> or something


Enumerations to implement:

* MSO_FILL_TYPE


* table cell fill color


Enumerations
------------

**MsoFillType**

http://msdn.microsoft.com/EN-US/library/office/ff861408.aspx

*msoFillBackground*
    5 -- Fill is the same as the background.

*msoFillGradient*
    3 -- Gradient fill.

*msoFillMixed*
    -2 -- Mixed fill.

*msoFillPatterned*
    2 -- Patterned fill.

*msoFillPicture*
    6 -- Picture fill.

*msoFillSolid*
    1 -- Solid fill.

*msoFillTextured*
    4 -- Textured fill.


XML specimens
-------------

.. highlight:: xml

Inherited fill on autoshape::

    <p:spPr>
       ...
      <a:prstGeom prst="roundRect">
        <a:avLst/>
      </a:prstGeom>
    </p:spPr>


Solid RGB color on autoshape::

    <p:spPr>
       ...
      <a:prstGeom prst="roundRect">
        <a:avLst/>
      </a:prstGeom>
      <a:solidFill>
        <a:srgbClr val="2CB731"/>
      </a:solidFill>
    </p:spPr>

Patterned fill::

    <a:pattFill prst="ltDnDiag">
      <a:fgClr>
        <a:schemeClr val="accent1"/>
      </a:fgClr>
      <a:bgClr>
        <a:schemeClr val="accent6"/>
      </a:bgClr>
    </a:pattFill>


XML semantics
-------------

* **No `prst` attribute.** When an `a:pattFill` element contains
  no `prst` attribute, the pattern default to 5% (dotted). This is the first
  one in the list on the PowerPoint UI.

* **No `fgClr` or `bgClr` elements.** When an `a:pattFill` element contains
  no `fgClr` or `bgClr` chile elements, the colors default to black and white
  respectively.



Resources
---------

* `MSDN FillFormat Object`_
* `MSDN MsoFillType Enumeration`_


.. _`MSDN FillFormat Object`:
   http://msdn.microsoft.com/en-us/library/office/ff744967.aspx

.. _`MSDN MsoFillType Enumeration`:
   http://msdn.microsoft.com/EN-US/library/office/ff861408.aspx

Related Schema Definitions
--------------------------

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
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode"/>
  </xsd:complexType>

  <xsd:group name="EG_Geometry">
    <xsd:choice>
      <xsd:element name="custGeom" type="CT_CustomGeometry2D"/>
      <xsd:element name="prstGeom" type="CT_PresetGeometry2D"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_FillProperties">
    <xsd:choice>
      <xsd:element name="noFill"    type="CT_NoFillProperties"/>
      <xsd:element name="solidFill" type="CT_SolidColorFillProperties"/>
      <xsd:element name="gradFill"  type="CT_GradientFillProperties"/>
      <xsd:element name="blipFill"  type="CT_BlipFillProperties"/>
      <xsd:element name="pattFill"  type="CT_PatternFillProperties"/>
      <xsd:element name="grpFill"   type="CT_GroupFillProperties"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_EffectProperties">
    <xsd:choice>
      <xsd:element name="effectLst" type="CT_EffectList"/>
      <xsd:element name="effectDag" type="CT_EffectContainer"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_BlipFillProperties">
    <xsd:sequence>
      <xsd:element name="blip"    type="CT_Blip"         minOccurs="0"/>
      <xsd:element name="srcRect" type="CT_RelativeRect" minOccurs="0"/>
      <xsd:group   ref="EG_FillModeProperties"           minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="dpi"          type="xsd:unsignedInt"/>
    <xsd:attribute name="rotWithShape" type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_GradientFillProperties">
    <xsd:sequence>
      <xsd:element name="gsLst"    type="CT_GradientStopList" minOccurs="0"/>
      <xsd:group   ref="EG_ShadeProperties"                   minOccurs="0"/>
      <xsd:element name="tileRect" type="CT_RelativeRect"     minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="flip"         type="ST_TileFlipMode"/>
    <xsd:attribute name="rotWithShape" type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_GroupFillProperties"/>

  <xsd:complexType name="CT_NoFillProperties"/>

  <xsd:complexType name="CT_PatternFillProperties">
    <xsd:sequence>
      <xsd:element name="fgClr" type="CT_Color" minOccurs="0"/>
      <xsd:element name="bgClr" type="CT_Color" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="prst" type="ST_PresetPatternVal"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Color">
    <xsd:sequence>
      <xsd:group ref="EG_ColorChoice"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SolidColorFillProperties">
    <xsd:sequence>
      <xsd:group ref="EG_ColorChoice" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="EG_ColorChoice">
    <xsd:choice>
      <xsd:element name="scrgbClr"  type="CT_ScRgbColor"/>
      <xsd:element name="srgbClr"   type="CT_SRgbColor"/>
      <xsd:element name="hslClr"    type="CT_HslColor"/>
      <xsd:element name="sysClr"    type="CT_SystemColor"/>
      <xsd:element name="schemeClr" type="CT_SchemeColor"/>
      <xsd:element name="prstClr"   type="CT_PresetColor"/>
    </xsd:choice>
  </xsd:group>

  <xsd:simpleType name="ST_PresetPatternVal">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="pct5"/>
      <xsd:enumeration value="pct10"/>
      <xsd:enumeration value="pct20"/>
      <xsd:enumeration value="pct25"/>
      <xsd:enumeration value="pct30"/>
      <xsd:enumeration value="pct40"/>
      <xsd:enumeration value="pct50"/>
      <xsd:enumeration value="pct60"/>
      <xsd:enumeration value="pct70"/>
      <xsd:enumeration value="pct75"/>
      <xsd:enumeration value="pct80"/>
      <xsd:enumeration value="pct90"/>
      <xsd:enumeration value="horz"/>
      <xsd:enumeration value="vert"/>
      <xsd:enumeration value="ltHorz"/>
      <xsd:enumeration value="ltVert"/>
      <xsd:enumeration value="dkHorz"/>
      <xsd:enumeration value="dkVert"/>
      <xsd:enumeration value="narHorz"/>
      <xsd:enumeration value="narVert"/>
      <xsd:enumeration value="dashHorz"/>
      <xsd:enumeration value="dashVert"/>
      <xsd:enumeration value="cross"/>
      <xsd:enumeration value="dnDiag"/>
      <xsd:enumeration value="upDiag"/>
      <xsd:enumeration value="ltDnDiag"/>
      <xsd:enumeration value="ltUpDiag"/>
      <xsd:enumeration value="dkDnDiag"/>
      <xsd:enumeration value="dkUpDiag"/>
      <xsd:enumeration value="wdDnDiag"/>
      <xsd:enumeration value="wdUpDiag"/>
      <xsd:enumeration value="dashDnDiag"/>
      <xsd:enumeration value="dashUpDiag"/>
      <xsd:enumeration value="diagCross"/>
      <xsd:enumeration value="smCheck"/>
      <xsd:enumeration value="lgCheck"/>
      <xsd:enumeration value="smGrid"/>
      <xsd:enumeration value="lgGrid"/>
      <xsd:enumeration value="dotGrid"/>
      <xsd:enumeration value="smConfetti"/>
      <xsd:enumeration value="lgConfetti"/>
      <xsd:enumeration value="horzBrick"/>
      <xsd:enumeration value="diagBrick"/>
      <xsd:enumeration value="solidDmnd"/>
      <xsd:enumeration value="openDmnd"/>
      <xsd:enumeration value="dotDmnd"/>
      <xsd:enumeration value="plaid"/>
      <xsd:enumeration value="sphere"/>
      <xsd:enumeration value="weave"/>
      <xsd:enumeration value="divot"/>
      <xsd:enumeration value="shingle"/>
      <xsd:enumeration value="wave"/>
      <xsd:enumeration value="trellis"/>
      <xsd:enumeration value="zigZag"/>
    </xsd:restriction>
  </xsd:simpleType>
