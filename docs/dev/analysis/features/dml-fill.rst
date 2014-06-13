
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

    fill = sp.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0x01, 0x23, 0x45)
    fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    fill.fore_color.brightness = 0.25
    fill.transparency = 0.25
    sp.fill = None
    fill.background()  # is almost free once the rest is in place


Candidate API
-------------

Protocol::

    >>> fill = sp.fill
    >>> assert(isinstance(fill, FillFormat)

    >>> fill.type
    None
    >>> fill.solid()
    >>> fill.type
    1  # MSO_FILL.SOLID

    >>> fill.fore_color = 'anything'
    AttributeError: can't set attribute  # .fore_color is read-only
    >>> fore_color = fill.fore_color
    >>> assert(isinstance(fore_color, ColorFormat))

    >>> fore_color.rgb = RGB(0x3F, 0x2c, 0x36)
    >>> fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    >>> fore_color.brightness = -0.25

    >>> fill.transparency
    0.0
    >>> fill.transparency = 0.25  # sets opacity to 75%

    >>> sp.fill = None  # removes any fill, fill is inherited from theme


API

* ``Shape.fill`` (instance of ``FillFormat``)


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


Resources
---------

* `MSDN FillFormat Object`_
* `MSDN MsoFillType Enumeration`_


.. _`MSDN FillFormat Object`:
   http://msdn.microsoft.com/en-us/library/office/ff744967.aspx

.. _`MSDN MsoFillType Enumeration`:
   http://msdn.microsoft.com/EN-US/library/office/ff861408.aspx
