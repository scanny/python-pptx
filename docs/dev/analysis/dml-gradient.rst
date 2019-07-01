.. _gradientfill:

Gradient Fill
=============

A gradient fill is a smooth gradual transition between two or more colors.
A gradient fill can mimic the shading of real-life objects and allow
a graphic to look less "flat".

Each "anchor" color and its relative location in the overall transition is
known as a *stop*. A gradient definition gathers these in a sequence known as
the *stop list*, which contains two or more stops (a gradient with only one
stop would be a solid color).

A gradient also has a *path*, which specifies the direction of the color
transition. Perhaps most common is a *linear path*, where the transition is
along a straight line, perhaps top-left to bottom-right. Other gradients like
*radial* and *rectangular* are possible. The initial implementation supports
only a linear path.


Protocol
--------

**Staring with a newly-created shape** A newly created shape has a fill of
|None|, indicating it's fill is inherited::

    >>> shape = shapes.add_shape(...)
    >>> fill = shape.fill
    >>> fill
    <pptx.dml.fill.FillFormat object at 0x...>
    >>> fill.type
    None

**Apply a gradient fill** A gradient fill is applied by calling the
`.gradient()` method on the `FillFormat` object. This applies a two-stop
linear gradient in the bottom-to-top direction (90-degrees in UI)::

    >>> fill.gradient()
    >>> fill._fill.xml
    <a:gradFill rotWithShape="1">
      <a:gsLst>
        <a:gs pos="0">
          <a:schemeClr val="accent1">
            <a:tint val="100000"/>
            <a:shade val="100000"/>
            <a:satMod val="130000"/>
          </a:schemeClr>
        </a:gs>
        <a:gs pos="100000">
          <a:schemeClr val="accent1">
            <a:tint val="50000"/>
            <a:shade val="100000"/>
            <a:satMod val="350000"/>
          </a:schemeClr>
        </a:gs>
      </a:gsLst>
      <a:lin ang="16200000" scaled="0"/>
    </a:gradFill>

**Change angle of linear gradient.**::

    >>> fill.gradient_angle
    270.0
    >>> fill.gradient_angle = 45
    >>> fill.gradient_angle
    45.0

**Access a stop.** The `GradientStops` sequence provides access to an
individual gradient stop.::

    >>> gradient_stops = fill.gradient_stops
    >>> gradient_stops
    <pptx.dml.fill.GradientStops object at 0x...>

    >>> len(gradient_stops)
    3

    >>> gradient_stop = gradient_stops[0]
    >>> gradient_stop
    <pptx.dml.fill.GradientStop object at 0x...>

**Manipulate gradient stop color.** The `.color` property of a gradient stop
is a `ColorFormat` object, which may be manipulated like any other color to
achieve the desired effect::

    >>> color = gradient_stop.color
    >>> color
    <pptx.dml.color.ColorFormat object at 0x...>
    >>> color.theme_color = MSO_THEME_COLOR.ACCENT_2

**Manipulate gradient stop position.**::

    >>> gradient_stop.position
    0.8  # ---represents 80%---
    >>> gradient_stop.position = 0.5
    >>> gradient_stop.position
    0.5

**Remove a gradient stop position.**::

    >>> gradient_stops[1].remove()
    >>> len(gradient_stops)
    2


PowerPoint UI features and behaviors
------------------------------------

* Gradient dialog enforces a two-stop minimum in the stops list. A stop that
  is one of only two stops cannot be deleted.

* Gradient Gallery ...


Related items from Microsoft VBA API
------------------------------------

* `GradientAngle`_
* `GradientColorType`_
* `GradientDegree`_
* `GradientStops`_
* `GradientStyle`_
* `GradientVariant`_

* `OneColorGradient()`_

* `PresetGradient()`_

* `TwoColorGradient()`_

.. _`GradientAngle`:
   https://msdn.microsoft.com/en-us/VBA/PowerPoint-VBA/articles/fillformat-gradientangle-property-powerpoint

.. _`GradientColorType`:
   https://msdn.microsoft.com/en-us/VBA/PowerPoint-VBA/articles/fillformat-gradientcolortype-property-powerpoint

.. _`GradientDegree`:
   https://msdn.microsoft.com/en-us/VBA/PowerPoint-VBA/articles/fillformat-gradientdegree-property-powerpoint

.. _`GradientStops`:
   https://msdn.microsoft.com/en-us/VBA/PowerPoint-VBA/articles/fillformat-gradientstops-property-powerpoint

.. _`GradientStyle`:
   https://msdn.microsoft.com/en-us/VBA/PowerPoint-VBA/articles/fillformat-gradientstyle-property-powerpoint

.. _`GradientVariant`:
   https://msdn.microsoft.com/en-us/VBA/PowerPoint-VBA/articles/fillformat-gradientvariant-property-powerpoint

.. _`OneColorGradient()`:
   https://msdn.microsoft.com/en-us/VBA/PowerPoint-VBA/articles/fillformat-p
   resetgradient-method-powerpoint

.. _`PresetGradient()`:
   https://msdn.microsoft.com/en-us/VBA/PowerPoint-VBA/articles/fillformat-o
   necolorgradient-method-powerpoint

.. _`TwoColorGradient()`:
   https://msdn.microsoft.com/en-us/VBA/PowerPoint-VBA/articles/fillformat-p
   resetgradient-method-powerpoint


Enumerations
------------

MsoFillType
~~~~~~~~~~~

http://msdn.microsoft.com/EN-US/library/office/ff861408.aspx

**msoFillBackground**
    5 -- Fill is the same as the background.

**msoFillGradient**
    3 -- Gradient fill.

**msoFillPatterned**
    2 -- Patterned fill.

**msoFillPicture**
    6 -- Picture fill.

**msoFillSolid**
    1 -- Solid fill.

**msoFillTextured**
    4 -- Textured fill.

**msoFillMixed**
    -2 -- Mixed fill.


MsoGradientStyle
~~~~~~~~~~~~~~~~

https://msdn.microsoft.com/en-us/vba/office-shared-vba/articles/msogradient\
style-enumeration-office

**msoGradientDiagonalDown**
    4    Diagonal gradient moving from a top corner down to the opposite corner.

**msoGradientDiagonalUp**
    3    Diagonal gradient moving from a bottom corner up to the opposite corner.

**msoGradientFromCenter**
    7    Gradient running from the center out to the corners.

**msoGradientFromCorner**
    5    Gradient running from a corner to the other three corners.

**msoGradientFromTitle**
    6    Gradient running from the title outward.

**msoGradientHorizontal**
    1    Gradient running horizontally across the shape.

**msoGradientVertical**
    2    Gradient running vertically down the shape.

**msoGradientMixed**
    \-2    Gradient is mixed.


MsoPresetGradientType
~~~~~~~~~~~~~~~~~~~~~

**msoGradientBrass**
    20    Brass gradient.

**msoGradientCalmWater**
    8    Calm Water gradient.

**msoGradientChrome**
    21    Chrome gradient.

**msoGradientChromeII**
    22    Chrome II gradient.

**msoGradientDaybreak**
    4    Daybreak gradient.

**msoGradientDesert**
    6    Desert gradient.

**msoGradientEarlySunset**
    1    Early Sunset gradient.

**msoGradientFire**
    9    Fire gradient.

**msoGradientFog**
    10    Fog gradient.

**msoGradientGold**
    18    Gold gradient.

**msoGradientGoldII**
    19    Gold II gradient.

**msoGradientHorizon**
    5    Horizon gradient.

**msoGradientLateSunset**
    2    Late Sunset gradient.

**msoGradientMahogany**
    15    Mahogany gradient.

**msoGradientMoss**
    11         Moss gradient.

**msoGradientNightfall**
    3    Nightfall gradient.

**msoGradientOcean**
    7         Ocean gradient.

**msoGradientParchment**
    14    Parchment gradient.

**msoGradientPeacock**
    12    Peacock gradient.

**msoGradientRainbow**
    16    Rainbow gradient.

**msoGradientRainbowII**
    17    Rainbow II gradient.

**msoGradientSapphire**
    24    Sapphire gradient.

**msoGradientSilver**
    23    Silver gradient.

**msoGradientWheat**
    13    Wheat gradient.

**msoPresetGradientMixed**
    -2    Mixed gradient.


XML specimens
-------------

.. highlight:: xml

Gradient fill (preset selected from gallery on PowerPoint 2014)::

    <a:gradFill flip="none" rotWithShape="1">
      <a:gsLst>
        <a:gs pos="0">
          <a:schemeClr val="accent1">
            <a:tint val="66000"/>
            <a:satMod val="160000"/>
          </a:schemeClr>
        </a:gs>
        <a:gs pos="50000">
          <a:schemeClr val="accent1">
            <a:tint val="44500"/>
            <a:satMod val="160000"/>
          </a:schemeClr>
        </a:gs>
        <a:gs pos="100000">
          <a:schemeClr val="accent1">
            <a:tint val="23500"/>
            <a:satMod val="160000"/>
          </a:schemeClr>
        </a:gs>
      </a:gsLst>
      <a:lin ang="2700000" scaled="1"/>
      <a:tileRect/>
    </a:gradFill>

Gradient fill (simple created with gradient dialog)::

    <a:gradFill flip="none" rotWithShape="1">
      <a:gsLst>
        <a:gs pos="0">
          <a:schemeClr val="accent1">
            <a:shade val="51000"/>
            <a:satMod val="130000"/>
          </a:schemeClr>
        </a:gs>
        <a:gs pos="100000">
          <a:schemeClr val="accent1">
            <a:lumMod val="40000"/>
            <a:lumOff val="60000"/>
          </a:schemeClr>
        </a:gs>
      </a:gsLst>
      <a:lin ang="2700000" scaled="0"/>
      <a:tileRect/>
    </a:gradFill>


XML semantics
-------------

* Each `a:gs` element is a fill format

* `a:lin@ang` is angle in 1/60,000ths of a degree. Zero degrees is the vector
  (1, 0), pointing directly to the right. Degrees are measured
  **counter-clockwise** from that origin.


Related Schema Definitions
--------------------------

::

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"    type="CT_Transform2D"            minOccurs="0"/>
      <xsd:choice minOccurs="0">  <!--EG_Geometry-->
        <xsd:element name="custGeom" type="CT_CustomGeometry2D"/>
        <xsd:element name="prstGeom" type="CT_PresetGeometry2D"/>
      </xsd:choice>
      <xsd:choice minOccurs="0">  <!--EG_FillProperties-->
        <xsd:element name="noFill"    type="CT_NoFillProperties"/>
        <xsd:element name="solidFill" type="CT_SolidColorFillProperties"/>
        <xsd:element name="gradFill"  type="CT_GradientFillProperties"/>
        <xsd:element name="blipFill"  type="CT_BlipFillProperties"/>
        <xsd:element name="pattFill"  type="CT_PatternFillProperties"/>
        <xsd:element name="grpFill"   type="CT_GroupFillProperties"/>
      </xsd:choice>
      <xsd:element name="ln"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:group   ref="EG_EffectProperties"                       minOccurs="0"/>
      <xsd:element name="scene3d" type="CT_Scene3D"                minOccurs="0"/>
      <xsd:element name="sp3d"    type="CT_Shape3D"                minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode"/>
  </xsd:complexType>

  <xsd:complexType name="CT_GradientFillProperties">
    <xsd:sequence>
      <xsd:element name="gsLst"  type="CT_GradientStopList" minOccurs="0"/>
      <xsd:choice minOccurs="0">  <!-- EG_ShadeProperties -->
        <xsd:element name="lin"  type="CT_LinearShadeProperties"/>
        <xsd:element name="path" type="CT_PathShadeProperties"/>
      </xsd:choice>
      <xsd:element name="tileRect" type="CT_RelativeRect" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="flip"         type="ST_TileFlipMode"/>
    <xsd:attribute name="rotWithShape" type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_GradientStopList">
    <xsd:sequence>
      <xsd:element name="gs" type="CT_GradientStop" minOccurs="2" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GradientStop">
    <xsd:sequence>
      <xsd:choice>  <!-- EG_ColorChoice --->
        <xsd:element name="scrgbClr"  type="CT_ScRgbColor"/>
        <xsd:element name="srgbClr"   type="CT_SRgbColor"/>
        <xsd:element name="hslClr"    type="CT_HslColor"/>
        <xsd:element name="sysClr"    type="CT_SystemColor"/>
        <xsd:element name="schemeClr" type="CT_SchemeColor"/>
        <xsd:element name="prstClr"   type="CT_PresetColor"/>
      </xsd:choice>
    </xsd:sequence>
    <xsd:attribute name="pos" type="ST_PositiveFixedPercentage" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_LinearShadeProperties">
    <xsd:attribute name="ang"    type="ST_PositiveFixedAngle"/>
    <xsd:attribute name="scaled" type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PathShadeProperties">
    <xsd:sequence>
      <xsd:element name="fillToRect" type="CT_RelativeRect" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="path" type="ST_PathShadeType" use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Color">
    <xsd:sequence>
      <xsd:group ref="EG_ColorChoice"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_RelativeRect">
    <xsd:attribute name="l" type="ST_Percentage" default="0%"/>
    <xsd:attribute name="t" type="ST_Percentage" default="0%"/>
    <xsd:attribute name="r" type="ST_Percentage" default="0%"/>
    <xsd:attribute name="b" type="ST_Percentage" default="0%"/>
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

  <xsd:simpleType name="ST_PathShadeType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="shape"/>
      <xsd:enumeration value="circle"/>
      <xsd:enumeration value="rect"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_PositiveFixedAngle">
    <xsd:restriction base="ST_Angle">
      <xsd:minInclusive value="0"/>
      <xsd:maxExclusive value="21600000"/>
    </xsd:restriction>

  <xsd:simpleType name="ST_PositiveFixedPercentage">
    <xsd:union memberTypes="
      ST_PositiveFixedPercentageDecimal
      s:ST_PositiveFixedPercentage
    "/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_PositiveFixedPercentageDecimal">
    <xsd:restriction base="ST_PercentageDecimal">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="100000"/>
    </xsd:restriction>
  </xsd:simpleType>

  <!-- s:ST_PositiveFixedPercentage -->
  <xsd:simpleType name="ST_PositiveFixedPercentage">
    <xsd:restriction base="ST_Percentage">
      <xsd:pattern value="((100)|([0-9][0-9]?))(\.[0-9][0-9]?)?%"/>
    </xsd:restriction>
  </xsd:simpleType>
