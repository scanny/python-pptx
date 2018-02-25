.. _ShapeShadow:

Shadow
======

Shadow is inherited, and a shape often appears with a shadow without
explicitly applying a shadow format.

The only difference in a shape with shadow turned off is the presence of an
empty `<a:effectLst/>` child in its `<p:spPr>` element.

Inherited shadow is turned off when `p:spPr/a:effectLst` is present with no
`a:outerShdw` child element. Other effect child elements may be present.

A shadow is one type of "effect". The others are glow/soft-edges and reflection.

Shadow may be of type *outer* (perhaps most common), *inner*, or
*perspective*.

* [ ] ? What shapes can have a shadow?

  + AutoShape
  + Connector
  + Picture
  + Graphics Frame
  + Group Shape

* [ ] Adding shadow to a group shape adds that shadow to each component shape
      in the group.

There is a "new shape format" concept. This format determines what a new
shape looks like, but does not change the appearance of shapes already in
place.

*Theme Effects* are a thing here. They are Subtle, Moderate, and Intense.

There are 40 built-in theme effects. Each of these have ...

Theme sub-tree `a:theme/a:objectDefaults/a:spDef/a:style/a:effectRef/idx=2`
specifies that new objects will get the second effect in
`a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst`. That effect looks
like this::

  <a:effectStyle>
    <a:effectLst>
      <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
        <a:srgbClr val="000000">
          <a:alpha val="35000"/>
        </a:srgbClr>
      </a:outerShdw>
    </a:effectLst>
  </a:effectStyle>


Scope
-----

* `Shape.shadow` is a `ShadowFormat` object (all shape types)

* `shape.shadow.inherit = True` removes any explicitly applied effects
  (including glow/soft-edge and reflection, not just shadow). This restores
  the "inherit" state for effects and makes the shape sensitive to changes in
  theme.

* `shape.shadow.visible = False` overrides any default (theme) effects. Any
  inherited shadow, glow, and/or reflection will no longer appear.

**Out-of-scope**

Minimum for specifying a basic shadow

* `ShadowFormat.shadow_type` (inner, outer, perspective)
* `ShadowFormat.alignment` (shadow anchor, automatic based on angle)
* `ShadowFormat.angle` (0-degrees is to the right, increasing CCW)
* `ShadowFormat.blur_radius` (generally a few points, maybe 3-5)
* `ShadowFormat.color` (theme or RGB color)
* `ColorFormat.transparency` (needed for proper shadow rendering)
* `ShadowFormat.distance` (generally a couple points, maybe 1-3)
* `shape.shadow.style` (MSO_SHADOW_STYLE) indicates whether shadow is inner,
  outer, or perspective (or None).

Nice to have for finer tuning

* `ShadowFormat.rotate_with_shape` (boolean)
* `ShadowFormat.scale` (shadow bigger/smaller than shape, default 100%)

* Retaining non-shadow effects when turning off shadow. This requires the
  ability to "clone" the currently enabled defaults. Effects are not
  inherited separately; making explicit the currently active default is how
  PowerPoint works around the "all-or-nothing" inheritance behavior.

* Clone effective shadow, glow, and/or reflection.


Protocol
--------

Assigning `None` removes any local override, causing shadow to be inherited
from parent. Note that it does not cause `shape.shadow` to return None;
`shape.shadow` *always* returns *the* `ShadowFormat` object for that shape::

    >>> shape.shadow
    <pptx.shapes.base.ShadowFormat object at 0x02468ac>
    >>> shape.shadow = None # ---shadow now appears as-inherited---
    >>> shape.shadow
    <pptx.shapes.base.ShadowFormat object at 0x02468ac>

To turn off an inherited shadow::

    >>> shadow = shape.shadow
    >>> shadow.visible = False

This has the side-effect of disabling inheritance of all effects for that
shape.


Specimen XML
------------

Shape inheriting shadow. Note the absence of `p:spPr/a:effectLst`, causing
all effects to be inherited::

      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Rounded Rectangle 3"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="4114800" y="2971800"/>
            <a:ext cx="914400" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="roundRect">
            <a:avLst/>
          </a:prstGeom>
        </p:spPr>
        <p:style>
          <a:lnRef idx="1">
            <a:schemeClr val="accent1"/>
          </a:lnRef>
          <a:fillRef idx="3">
            <a:schemeClr val="accent1"/>
          </a:fillRef>
          <a:effectRef idx="2">
            <a:schemeClr val="accent1"/>
          </a:effectRef>
          <a:fontRef idx="minor">
            <a:schemeClr val="lt1"/>
          </a:fontRef>
        </p:style>
        <p:txBody>
          <a:bodyPr rtlCol="0" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </p:txBody>
      </p:sp>

Shape with inherited shadow turned off::

      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Rounded Rectangle 3"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="4114800" y="2971800"/>
            <a:ext cx="914400" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="roundRect">
            <a:avLst/>
          </a:prstGeom>
          <a:effectLst/>
        </p:spPr>
        <p:style>
          <a:lnRef idx="1">
            <a:schemeClr val="accent1"/>
          </a:lnRef>
          <a:fillRef idx="3">
            <a:schemeClr val="accent1"/>
          </a:fillRef>
          <a:effectRef idx="2">
            <a:schemeClr val="accent1"/>
          </a:effectRef>
          <a:fontRef idx="minor">
            <a:schemeClr val="lt1"/>
          </a:fontRef>
        </p:style>
        <p:txBody>
          <a:bodyPr rtlCol="0" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </p:txBody>
      </p:sp>


XML Semantics
-------------

**Effect inheritance is "all-or-nothing"**

* If `p:spPr/a:effectLst` is present, all desired effects must be specified
  explicitly as its children; a missing child, such as `a:outerShdw`, will
  cause that effect to be turned off. PowerPoint automatically adds those
  populated with inherited values when one of the effects is customized,
  necessitating that addition of an `a:effectLst` element.


PowerPoint behaviors
--------------------

* Set visible off for a customized shadow removes all customized settings and
  they are not recoverable by setting the shadow visible again.


MS API
------

ShadowFormat object
~~~~~~~~~~~~~~~~~~~

* `ShadowFormat.Visible`


Schema excerpt
--------------

::

  <xsd:complexType name="CT_ShapeProperties">  <!--denormalized-->
    <xsd:sequence>
      <xsd:element name="xfrm"              type="CT_Transform2D"            minOccurs="0"/>
      <xsd:group   ref ="EG_Geometry"                                        minOccurs="0"/>
      <xsd:group   ref ="EG_FillProperties"                                  minOccurs="0"/>
      <xsd:element name="ln"                type="CT_LineProperties"         minOccurs="0"/>
      <xsd:choice minOccurs="0"/>  <!--EG_EffectProperties-->
        <xsd:element name="effectLst"       type="CT_EffectList"/>
        <xsd:element name="effectDag"       type="CT_EffectContainer"/>
      </xsd:choice>
      <xsd:element name="scene3d"           type="CT_Scene3D"                minOccurs="0"/>
      <xsd:element name="sp3d"              type="CT_Shape3D"                minOccurs="0"/>
      <xsd:element name="extLst"            type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode"/>
  </xsd:complexType>

  <xsd:complexType name="CT_EffectList">
    <xsd:sequence>
      <xsd:element name="blur"        type="CT_BlurEffect"         minOccurs="0"/>
      <xsd:element name="fillOverlay" type="CT_FillOverlayEffect"  minOccurs="0"/>
      <xsd:element name="glow"        type="CT_GlowEffect"         minOccurs="0"/>
      <xsd:element name="innerShdw"   type="CT_InnerShadowEffect"  minOccurs="0"/>
      <xsd:element name="outerShdw"   type="CT_OuterShadowEffect"  minOccurs="0"/>
      <xsd:element name="prstShdw"    type="CT_PresetShadowEffect" minOccurs="0"/>
      <xsd:element name="reflection"  type="CT_ReflectionEffect"   minOccurs="0"/>
      <xsd:element name="softEdge"    type="CT_SoftEdgesEffect"    minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_OuterShadowEffect">
    <xsd:sequence>
      <xsd:group ref="EG_ColorChoice" minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="blurRad"      type="ST_PositiveCoordinate" default="0"/>
    <xsd:attribute name="dist"         type="ST_PositiveCoordinate" default="0"/>
    <xsd:attribute name="dir"          type="ST_PositiveFixedAngle" default="0"/>
    <xsd:attribute name="sx"           type="ST_Percentage"         default="100%"/>
    <xsd:attribute name="sy"           type="ST_Percentage"         default="100%"/>
    <xsd:attribute name="kx"           type="ST_FixedAngle"         default="0"/>
    <xsd:attribute name="ky"           type="ST_FixedAngle"         default="0"/>
    <xsd:attribute name="algn"         type="ST_RectAlignment"      default="b"/>
    <xsd:attribute name="rotWithShape" type="xsd:boolean"           default="true"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_RectAlignment">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="tl"/>
      <xsd:enumeration value="t"/>
      <xsd:enumeration value="tr"/>
      <xsd:enumeration value="l"/>
      <xsd:enumeration value="ctr"/>
      <xsd:enumeration value="r"/>
      <xsd:enumeration value="bl"/>
      <xsd:enumeration value="b"/>
      <xsd:enumeration value="br"/>
    </xsd:restriction>
  </xsd:simpleType>
