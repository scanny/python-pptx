
Font Color
==========

Overview
--------

Font color is a particular case of the broader topic of *fill*, the case of
*solid fill*, in which a drawing element is filled with a single color. Other
possible fill types are *gradient*, *picture*, *pattern*, and *background*
(transparent).

Any drawing element that can have a fill can have a solid fill, or color.
Elements that can have a fill include shape, line, text, table, table cell, and
slide background.


Scope
-----

This analysis focuses on *font color*, although much of it is general to solid
fills. The focus is simply to enable getting to feature implementation as
quickly as possible, just not before its clear I understand how it works in
general.


Candidate API
-------------

New members:

* ``Font.color``
* ``ColorFormat(fillable_elm)``
* ``ColorFormat.type``
* ``ColorFormat.rgb``
* ``ColorFormat.theme_color``
* ``ColorFormat.brightness``
* ``RGBColor(r, g, b)``
* ``RGBColor.__str__``
* ``RGBColor.from_str(rgb_str)``

Enumerations:

* MSO_COLOR_TYPE_INDEX
* MSO_THEME_COLOR_INDEX


Protocol::

    >>> assert isinstance(font.color, ColorFormat)

    >>> font.color = 'anything'
    AttributeError: can't set attribute

    >>> assert font.color.type in MSO_COLOR_TYPE_INDEX
    wouldn't actually work, but conceptually true

    >>> font.color.rgb = RGB(0x3F, 0x2c, 0x36)

    >>> font.color.theme_color = MSO_THEME_COLOR.ACCENT_1

    >>> font.color.brightness = -0.25


XML specimens
-------------

Here is a representative sample of the various font color cases as they appear
in the XML, as produced by PowerPoint. Some inoperative attributes have been
removed for clarity.

Baseline run::

    <a:r>
      <a:rPr/>
      <a:t>test text</a:t>
    </a:r>

set to a specific RGB color, using color wheel; HSB sliders yield the same::

    <a:rPr>
      <a:solidFill>
        <a:srgbClr val="748E1D"/>
      </a:solidFill>
    </a:rPr>

set to theme color, Text 2::

    <a:rPr>
      <a:solidFill>
        <a:schemeClr val="tx2"/>
      </a:solidFill>
    </a:rPr>

set to theme color, Accent 2, 80% Lighter::

    <a:rPr>
      <a:solidFill>
        <a:schemeClr val="accent2">
          <a:lumMod val="20000"/>
          <a:lumOff val="80000"/>
        </a:schemeClr>
      </a:solidFill>
    </a:rPr>


PowerPoint API
--------------

Changing the color of text in the PowerPoint API is accomplished with something
like this::

    Set font = textbox.TextFrame.TextRange.Font

    'set font to a specific RGB color
    font.Color.RGB = RGB(12, 34, 56)

    Debug.Print font.Color.RGB
    > 3678732

    'set font to a theme color; Accent 1, 25% Darker in this example
    font.Color.ObjectThemeColor = msoThemeColorAccent1
    font.Color.Brightness = -0.25

The ColorFormat object is the first interesting object here, the type of object
returned by ``TextRange.Font.Color``. It includes the following properties that
will likely have counterparts in python-pptx:

Brightness
    Returns or sets the brightness of the specified object. The value for
    this property must be a number from -1.0 (darker) to 1.0 (lighter), with
    0 corresponding to no brightness adjustment. Read/write Single. Note: this
    corresponds to selecting an adjusted theme color from the PowerPoint ribbon
    color picker, like 'Accent 1, 40% Lighter'.

ObjectThemeColor
    Returns or sets the theme color of the specified ColorFormat object.
    Read/Write. Accepts and returns values from the enumeration
    MsoThemeColorIndex.

RGB
    Returns or sets the red-green-blue (RGB) value of the specified color.
    Read/write.

Type
    Represents the type of color. Read-only.


Legacy properties
~~~~~~~~~~~~~~~~~

These two properties will probably not need to be implemented in python-pptx.

SchemeColor
    Returns or sets the color in the applied color scheme that's associated
    with the specified object. Accepts and returns values from
    PpColorSchemeIndex. Read/write. Appears to be a legacy method to accomodate
    code prior to PowerPoint 2007.

TintAndShade
    Sets or returns the lightening or darkening of the the color of a specified
    shape. Read/write.


Resources
~~~~~~~~~

* `MSDN TextFrame2 Members`_
* `MSDN TextRange Members`_
* `MSDN Font Members`_
* `MSDN ColorFormat Members`_
* `MSDN MsoThemeColorIndex Enumeration`_


.. _`MSDN TextFrame2 Members`:
   http://msdn.microsoft.com/en-us/library/office/ff746114.aspx

.. _`MSDN TextRange Members`:
   http://msdn.microsoft.com/en-us/library/office/ff746274.aspx

.. _`MSDN Font Members`:
   http://msdn.microsoft.com/en-us/library/office/ff745818.aspx

.. _`MSDN ColorFormat Members`:
   http://msdn.microsoft.com/en-us/library/office/ff745051.aspx

.. _`MSDN MsoThemeColorIndex Enumeration`:
   http://msdn.microsoft.com/en-us/library/office/ff860782.aspx


Behaviors
---------

The API method ``Brightness`` corresponds to the UI action of selecting an
auto-generated tint or shade of a theme color from the PowerPoint ribbon color
picker:

Setting font color to Accent 1 from the UI produces::

    <a:rPr>
      <a:solidFill>
        <a:schemeClr val="accent1"/>
      </a:solidFill>
    </a:rPr>

The following code does the same from the API::

    fnt.Color.ObjectThemeColor = msoThemeColorAccent1

Setting font color to Accent 1, Lighter 40% (40% tint) from the PowerPoint UI
produces this XML::

    <a:rPr>
      <a:solidFill>
        <a:schemeClr val="accent1">
          <a:lumMod val="60000"/>
          <a:lumOff val="40000"/>
        </a:schemeClr>
      </a:solidFill>
    </a:rPr>

Setting ``Brightness`` to +0.4 has the same effect::

    fnt.Color.Brightness = 0.4

Setting font color to Accent 1, Darker 25% (25% shade) from the UI results in
the following XML. Note that no ``<a:lumOff>`` element is used.::

    <a:rPr>
      <a:solidFill>
        <a:schemeClr val="accent1">
          <a:lumMod val="75000"/>
        </a:schemeClr>
      </a:solidFill>
    </a:rPr>

Setting ``Brightness`` to -0.25 has the same effect::

    fnt.Color.Brightness = -0.25

Calling ``TintAndShade`` with a positive value (between 0 and 1) causes a tint
element to be inserted, but I'm not at all sure why and when one would want to
use it rather than the ``Brightness`` property.::

    fnt.Color.TintAndShade = 0.75

    <a:rPr>
      <a:solidFill>
        <a:schemeClr val="accent1">
          <a:tint val="25000"/>
        </a:schemeClr>
      </a:solidFill>
    </a:rPr>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_SolidColorFillProperties">
    <xsd:sequence>
      <xsd:group ref="EG_ColorChoice" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="EG_ColorChoice">
    <xsd:choice>
      <xsd:element name="scrgbClr"  type="CT_ScRgbColor"  minOccurs="1" maxOccurs="1"/>
      <xsd:element name="srgbClr"   type="CT_SRgbColor"   minOccurs="1" maxOccurs="1"/>
      <xsd:element name="hslClr"    type="CT_HslColor"    minOccurs="1" maxOccurs="1"/>
      <xsd:element name="sysClr"    type="CT_SystemColor" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="schemeClr" type="CT_SchemeColor" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="prstClr"   type="CT_PresetColor" minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_SRgbColor">
    <xsd:sequence>
      <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="val" type="s:ST_HexColorRGB" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_SchemeColor">
    <xsd:sequence>
      <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="val" type="ST_SchemeColorVal" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PresetColor">
    <xsd:sequence>
      <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="val" type="ST_PresetColorVal" use="required"/>
  </xsd:complexType>

  <xsd:group name="EG_ColorTransform">
    <xsd:choice>
      <xsd:element name="tint"     type="CT_PositiveFixedPercentage" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="shade"    type="CT_PositiveFixedPercentage" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="comp"     type="CT_ComplementTransform"     minOccurs="1" maxOccurs="1"/>
      <xsd:element name="inv"      type="CT_InverseTransform"        minOccurs="1" maxOccurs="1"/>
      <xsd:element name="gray"     type="CT_GrayscaleTransform"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="alpha"    type="CT_PositiveFixedPercentage" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="alphaOff" type="CT_FixedPercentage"         minOccurs="1" maxOccurs="1"/>
      <xsd:element name="alphaMod" type="CT_PositivePercentage"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="hue"      type="CT_PositiveFixedAngle"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="hueOff"   type="CT_Angle"                   minOccurs="1" maxOccurs="1"/>
      <xsd:element name="hueMod"   type="CT_PositivePercentage"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="sat"      type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="satOff"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="satMod"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="lum"      type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="lumOff"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="lumMod"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="red"      type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="redOff"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="redMod"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="green"    type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="greenOff" type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="greenMod" type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="blue"     type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="blueOff"  type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="blueMod"  type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="gamma"    type="CT_GammaTransform"          minOccurs="1" maxOccurs="1"/>
      <xsd:element name="invGamma" type="CT_InverseGammaTransform"   minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_Percentage">
    <xsd:attribute name="val" type="ST_Percentage" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Percentage">
    <xsd:union memberTypes="s:ST_Percentage"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Percentage">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="-?[0-9]+(\.[0-9]+)?%"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_HexColorRGB">
    <xsd:restriction base="xsd:hexBinary">
      <xsd:length value="3" fixed="true"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_SchemeColorVal">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="bg1"/>
      <xsd:enumeration value="tx1"/>
      <xsd:enumeration value="bg2"/>
      <xsd:enumeration value="tx2"/>
      <xsd:enumeration value="accent1"/>
      <xsd:enumeration value="accent2"/>
      <xsd:enumeration value="accent3"/>
      <xsd:enumeration value="accent4"/>
      <xsd:enumeration value="accent5"/>
      <xsd:enumeration value="accent6"/>
      <xsd:enumeration value="hlink"/>
      <xsd:enumeration value="folHlink"/>
      <xsd:enumeration value="phClr"/>
      <xsd:enumeration value="dk1"/>
      <xsd:enumeration value="lt1"/>
      <xsd:enumeration value="dk2"/>
      <xsd:enumeration value="lt2"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:group name="EG_FillProperties">
    <xsd:choice>
      <xsd:element name="noFill"    type="CT_NoFillProperties"         minOccurs="1" maxOccurs="1"/>
      <xsd:element name="solidFill" type="CT_SolidColorFillProperties" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="gradFill"  type="CT_GradientFillProperties"   minOccurs="1" maxOccurs="1"/>
      <xsd:element name="blipFill"  type="CT_BlipFillProperties"       minOccurs="1" maxOccurs="1"/>
      <xsd:element name="pattFill"  type="CT_PatternFillProperties"    minOccurs="1" maxOccurs="1"/>
      <xsd:element name="grpFill"   type="CT_GroupFillProperties"      minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_NoFillProperties"/>

  <xsd:complexType name="CT_GradientFillProperties">
    <xsd:sequence>
      <xsd:element name="gsLst"    type="CT_GradientStopList" minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_ShadeProperties"                   minOccurs="0" maxOccurs="1"/>
      <xsd:element name="tileRect" type="CT_RelativeRect"     minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="flip"         type="ST_TileFlipMode" use="optional"/>
    <xsd:attribute name="rotWithShape" type="xsd:boolean"     use="optional"/>
  </xsd:complexType>


Enumerations
------------

**MsoColorType**

http://msdn.microsoft.com/en-us/library/office/aa432491(v=office.12).aspx

msoColorTypeRGB
    1 - Color is determined by values of red, green, and blue.

msoColorTypeScheme
    2 - Color is defined by an application-specific scheme.

msoColorTypeCMYK
    3 - Color is determined by values of cyan, magenta, yellow, and black.

msoColorTypeCMS
    4 - Color Management System color type.

msoColorTypeInk
    5 - Not supported.

msoColorTypeMixed
    -2 - Not supported.

**MsoThemeColorIndex**

http://msdn.microsoft.com/en-us/library/office/aa432702(v=office.12).aspx

msoNotThemeColor
    0 - Specifies no theme color.
msoThemeColorDark1
    1 - Specifies the Dark 1 theme color.
msoThemeColorLight1
    2 - Specifies the Light 1 theme color.
msoThemeColorDark2
    3 - Specifies the Dark 2 theme color.
msoThemeColorLight2
    4 - Specifies the Light 2 theme color.
msoThemeColorAccent1
    5 - Specifies the Accent 1 theme color.
msoThemeColorAccent2
    6 - Specifies the Accent 2 theme color.
msoThemeColorAccent3
    7 - Specifies the Accent 3 theme color.
msoThemeColorAccent4
    8 - Specifies the Accent 4 theme color.
msoThemeColorAccent5
    9 - Specifies the Accent 5 theme color.
msoThemeColorAccent6
    10 - Specifies the Accent 6 theme color.
msoThemeColorHyperlink
    11 - Specifies the theme color for a hyperlink.
msoThemeColorFollowedHyperlink
    12 - Specifies the theme color for a clicked hyperlink.
msoThemeColorText1
    13 - Specifies the Text 1 theme color.
msoThemeColorBackground1
    14 - Specifies the Background 1 theme color.
msoThemeColorText2
    15 - Specifies the Text 2 theme color.
msoThemeColorBackground2
    16 - Specifies the Background 2 theme color.
msoThemeColorMixed
    -2 - Specifies a mixed color theme.


Value Objects
-------------

RGB
    RGBColor would be an immutable value object that could be reused as often
    as needed and not tied to any part of the underlying XML tree.


Other possible bits
-------------------

* acceptance test sketch
* test data requirements; files, builder(s)
* enumerations and mappings
* value types required
* test criteria

Example test criteria::

   # XML
   <a:ln>
     <a:solidFill>
       <a:srgbClr val="123456"/>
     </a:solidFill>
   </a:ln>

   assert font.color.type == MSO_COLOR_TYPE.RGB
   assert font.color.rgb == RGB(0x12, 0x34, 0x56)
   assert font.color.schemeClr == MSO_THEME_COLOR.NONE
   assert font.color.brightness == 0.0
