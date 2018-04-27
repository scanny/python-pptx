.. _SlideBackground:

Slide Background
================

A slide inherits its background from its layout or master (and possibly
theme), in that order. Each slide can also override its inherited background
settings.

A background is essentially a container for a fill.


Protocol
--------

**Access slide background.** `Slide.background` is the `Background` object
for the slide. Its presence is unconditional and the same `Background` object
is provided on each access for a particular `Slide` instance. Its presence
does not imply an explicitly applied slide background::

    >>> background = slide.background
    >>> background
    <pptx.shared.Background object at 0x0...>

**Detect explicitly-applied slide background.** A slide inherits its
background unless an explicitly-defined background has been applied to the
slide. The presence of an "override" background is determined by
interrogating the `.follow_master_background` property on each slide::

    >>> slide.follow_master_background
    True

**Override inheritance of background.** Assigning `False` to
`.follow_master_background` adds a "blank" background to the slide, which
then no longer inherits any layout or master background::

    >>> slide.follow_master_background = False
    >>> slide.follow_master_background
    False

Note that this step is not necessary to override inheritance. Merely
specifying a background fill will also interrupt inheritance. This approach
might be the easiest way though, if all you want is to interrupt inheritance
and don't want to apply a particular fill.

**Restore inheritance of slide background.** Any explicitly-applied slide
background can be removed by assigning `True` to
`.follow_master_background`::

    >>> slide.follow_master_background = True
    >>> slide.follow_master_background
    True

**Access background fill.** The `FillFormat` object for a slide background is
accessed using the background's `.fill` property. Note that merely accessing
this property will suppress inheritance of background for this slide. This
shouldn't normally be a problem as there would be little reason to access the
property without intention to change it::

    >>> background.fill
    <pptx.dml.fill.FillFormat object at 0x0...>

**Apply solid color background.** A background color is specified in the same
way as fill is specified for a shape. Note that the `FillFormat` object also
supports applying theme colors and patterns::

    >>> fill = background.fill
    >>> fill.solid()
    >>> fill.fore_color.rgb = RGBColor(255, 0, 0)

Minimal `p:bg`?::

  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:noFill/>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    ...
  </p:cSld>


Schema excerpt
--------------

.. highlight:: xml

::

  <xsd:element name="sld" type="CT_Slide"/>

  <xsd:complexType name="CT_Slide">  <!-- denormalized -->
    <xsd:sequence minOccurs="1" maxOccurs="1">
      <xsd:element name="cSld"       type="CT_CommonSlideData"/>
      <xsd:element name="clrMapOvr"  type="a:CT_ColorMappingOverride" minOccurs="0"/>
      <xsd:element name="transition" type="CT_SlideTransition"        minOccurs="0"/>
      <xsd:element name="timing"     type="CT_SlideTiming"            minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_ExtensionListModify"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="showMasterSp"     type="xsd:boolean" default="true"/>
    <xsd:attribute name="showMasterPhAnim" type="xsd:boolean" default="true"/>
    <xsd:attribute name="show"             type="xsd:boolean" default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_CommonSlideData">
    <xsd:sequence>
      <xsd:element name="bg"          type="CT_Background"       minOccurs="0"/>
      <xsd:element name="spTree"      type="CT_GroupShape"/>
      <xsd:element name="custDataLst" type="CT_CustomerDataList" minOccurs="0"/>
      <xsd:element name="controls"    type="CT_ControlList"      minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="name" type="xsd:string" use="optional" default=""/>
  </xsd:complexType>

  <xsd:complexType name="CT_Background">
    <xsd:choice>  <!-- EG_Background - one and only one -->
      <xsd:element name="bgPr"  type="CT_BackgroundProperties"/>
      <xsd:element name="bgRef" type="a:CT_StyleMatrixReference"/>
    </xsd:choice>
    <xsd:attribute name="bwMode" type="a:ST_BlackWhiteMode" use="optional" default="white"/>
  </xsd:complexType>

  <xsd:complexType name="CT_BackgroundProperties">
    <xsd:sequence>
      <xsd:group ref="a:EG_FillProperties"/>
      <xsd:group ref="a:EG_EffectProperties" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="shadeToTitle" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

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

  <xsd:simpleType name="ST_SlideId">
    <xsd:restriction base="xsd:unsignedInt">
      <xsd:minInclusive value="256"/>
      <xsd:maxExclusive value="2147483648"/>
    </xsd:restriction>
  </xsd:simpleType>

