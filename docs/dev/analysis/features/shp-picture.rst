==============
``CT_Picture``
==============

.. highlight:: xml

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 15, 50

   Schema Name  , CT_Picture
   Spec Name    , Picture
   Tag(s)       , p:pic
   Namespace    , presentationml (pml.xsd)
   Schema Line  , 1245
   Spec Section , 19.3.1.37


Example
=======

::

  <p:pic>
    <p:nvPicPr>
      <p:cNvPr id="6" name="Picture 5" descr="python-logo.gif"/>
      <p:cNvPicPr>
        <a:picLocks noChangeAspect="1"/>
      </p:cNvPicPr>
      <p:nvPr/>
    </p:nvPicPr>
    <p:blipFill>
      <a:blip r:embed="rId2"/>
      <a:stretch>
        <a:fillRect/>
      </a:stretch>
    </p:blipFill>
    <p:spPr>
      <a:xfrm>
        <a:off x="5580112" y="1988840"/>
        <a:ext cx="2679700" cy="901700"/>
      </a:xfrm>
      <a:prstGeom prst="rect">
        <a:avLst/>
      </a:prstGeom>
      <a:ln>
        <a:solidFill>
          <a:schemeClr val="bg1">
            <a:lumMod val="85000"/>
          </a:schemeClr>
        </a:solidFill>
      </a:ln>
    </p:spPr>
  </p:pic>


Minimal ``pic`` shape
=====================

::

  <p:pic>
    <p:nvPicPr>
      <p:cNvPr id="9" name="Picture 8" descr="python-logo.gif"/>
      <p:cNvPicPr/>
      <p:nvPr/>
    </p:nvPicPr>
    <p:blipFill>
      <a:blip r:embed="rId9"/>
    </p:blipFill>
    <p:spPr/>
  </p:pic>


Analysis
========

...


attributes
^^^^^^^^^^

None.


child elements
^^^^^^^^^^^^^^

=========  ===  =======================  ==========
name        #   type                     line
=========  ===  =======================  ==========
nvPicPr     1   CT_PictureNonVisual      1236
blipFill    1   a:CT_BlipFillProperties  1489 dml
spPr        1   a:CT_ShapeProperties     2210 dml
style       ?   a:CT_ShapeStyle          2245 dml
extLst      ?   CT_ExtensionListModify
=========  ===  =======================  ==========


Spec text
^^^^^^^^^

   This element specifies the existence of a picture object within the
   document.


Schema excerpt
^^^^^^^^^^^^^^

::

  <xsd:complexType name="CT_Picture">
    <xsd:sequence>
      <xsd:element name="nvPicPr"  type="CT_PictureNonVisual"     minOccurs="1" maxOccurs="1"/>
      <xsd:element name="blipFill" type="a:CT_BlipFillProperties" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="spPr"     type="a:CT_ShapeProperties"    minOccurs="1" maxOccurs="1"/>
      <xsd:element name="style"    type="a:CT_ShapeStyle"         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"   type="CT_ExtensionListModify"  minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_PictureNonVisual">
    <xsd:sequence>
      <xsd:element name="cNvPr"    type="a:CT_NonVisualDrawingProps"          minOccurs="1" maxOccurs="1"/>
      <xsd:element name="cNvPicPr" type="a:CT_NonVisualPictureProperties"     minOccurs="1" maxOccurs="1"/>
      <xsd:element name="nvPr"     type="CT_ApplicationNonVisualDrawingProps" minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="hlinkClick" type="CT_Hyperlink"              minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hlinkHover" type="CT_Hyperlink"              minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"     type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="id"     type="ST_DrawingElementId" use="required"/>
    <xsd:attribute name="name"   type="xsd:string"          use="required"/>
    <xsd:attribute name="descr"  type="xsd:string"          use="optional" default=""/>
    <xsd:attribute name="hidden" type="xsd:boolean"         use="optional" default="false"/>
    <xsd:attribute name="title"  type="xsd:string"          use="optional" default=""/>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualPictureProperties">
    <xsd:sequence>
      <xsd:element name="picLocks" type="CT_PictureLocking"         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"   type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="preferRelativeResize" type="xsd:boolean" use="optional" default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ApplicationNonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="ph"          type="CT_Placeholder"      minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="a:EG_Media"                              minOccurs="0" maxOccurs="1"/>
      <xsd:element name="custDataLst" type="CT_CustomerDataList" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"    minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="isPhoto"   type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="userDrawn" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:group name="EG_Media">
    <xsd:choice>
      <xsd:element name="audioCd"       type="CT_AudioCD"/>
      <xsd:element name="wavAudioFile"  type="CT_EmbeddedWAVAudioFile"/>
      <xsd:element name="audioFile"     type="CT_AudioFile"/>
      <xsd:element name="videoFile"     type="CT_VideoFile"/>
      <xsd:element name="quickTimeFile" type="CT_QuickTimeFile"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_BlipFillProperties">
    <xsd:sequence>
      <xsd:element name="blip"    type="CT_Blip"         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="srcRect" type="CT_RelativeRect" minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_FillModeProperties"           minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="dpi"          type="xsd:unsignedInt" use="optional"/>
    <xsd:attribute name="rotWithShape" type="xsd:boolean"     use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Blip">
    <xsd:sequence>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="alphaBiLevel" type="CT_AlphaBiLevelEffect"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaCeiling" type="CT_AlphaCeilingEffect"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaFloor"   type="CT_AlphaFloorEffect"         minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaInv"     type="CT_AlphaInverseEffect"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaMod"     type="CT_AlphaModulateEffect"      minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaModFix"  type="CT_AlphaModulateFixedEffect" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="alphaRepl"    type="CT_AlphaReplaceEffect"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="biLevel"      type="CT_BiLevelEffect"            minOccurs="1" maxOccurs="1"/>
        <xsd:element name="blur"         type="CT_BlurEffect"               minOccurs="1" maxOccurs="1"/>
        <xsd:element name="clrChange"    type="CT_ColorChangeEffect"        minOccurs="1" maxOccurs="1"/>
        <xsd:element name="clrRepl"      type="CT_ColorReplaceEffect"       minOccurs="1" maxOccurs="1"/>
        <xsd:element name="duotone"      type="CT_DuotoneEffect"            minOccurs="1" maxOccurs="1"/>
        <xsd:element name="fillOverlay"  type="CT_FillOverlayEffect"        minOccurs="1" maxOccurs="1"/>
        <xsd:element name="grayscl"      type="CT_GrayscaleEffect"          minOccurs="1" maxOccurs="1"/>
        <xsd:element name="hsl"          type="CT_HSLEffect"                minOccurs="1" maxOccurs="1"/>
        <xsd:element name="lum"          type="CT_LuminanceEffect"          minOccurs="1" maxOccurs="1"/>
        <xsd:element name="tint"         type="CT_TintEffect"               minOccurs="1" maxOccurs="1"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attributeGroup ref="AG_Blob"/>
    <xsd:attribute name="cstate" type="ST_BlipCompression" use="optional" default="none"/>
  </xsd:complexType>

  <xsd:attributeGroup name="AG_Blob">
    <xsd:attribute ref="r:embed" use="optional" default=""/>
    <xsd:attribute ref="r:link"  use="optional" default=""/>
  </xsd:attributeGroup>

  <xsd:group name="EG_FillModeProperties">
    <xsd:choice>
      <xsd:element name="tile"    type="CT_TileInfoProperties"    minOccurs="1" maxOccurs="1"/>
      <xsd:element name="stretch" type="CT_StretchInfoProperties" minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_StretchInfoProperties">
    <xsd:sequence>
      <xsd:element name="fillRect" type="CT_RelativeRect" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_RelativeRect">
    <xsd:attribute name="l" type="ST_Percentage" use="optional" default="0%"/>
    <xsd:attribute name="t" type="ST_Percentage" use="optional" default="0%"/>
    <xsd:attribute name="r" type="ST_Percentage" use="optional" default="0%"/>
    <xsd:attribute name="b" type="ST_Percentage" use="optional" default="0%"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"    type="CT_Transform2D"            minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_Geometry"                                 minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_FillProperties"                           minOccurs="0" maxOccurs="1"/>
      <xsd:element name="ln"      type="CT_LineProperties"         minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_EffectProperties"                         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="scene3d" type="CT_Scene3D"                minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sp3d"    type="CT_Shape3D"                minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Transform2D">
    <xsd:sequence>
      <xsd:element name="off" type="CT_Point2D" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="ext" type="CT_PositiveSize2D" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="rot" type="ST_Angle" use="optional" default="0"/>
    <xsd:attribute name="flipH" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="flipV" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Point2D">
    <xsd:attribute name="x" type="ST_Coordinate" use="required"/>
    <xsd:attribute name="y" type="ST_Coordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PositiveSize2D">
    <xsd:attribute name="cx" type="ST_PositiveCoordinate" use="required"/>
    <xsd:attribute name="cy" type="ST_PositiveCoordinate" use="required"/>
  </xsd:complexType>

  <xsd:group name="EG_Geometry">
    <xsd:choice>
      <xsd:element name="custGeom" type="CT_CustomGeometry2D" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="prstGeom" type="CT_PresetGeometry2D" minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

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

  <xsd:group name="EG_EffectProperties">
    <xsd:choice>
      <xsd:element name="effectLst" type="CT_EffectList"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="effectDag" type="CT_EffectContainer" minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_ShapeStyle">
    <xsd:sequence>
      <xsd:element name="lnRef"     type="CT_StyleMatrixReference" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="fillRef"   type="CT_StyleMatrixReference" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="effectRef" type="CT_StyleMatrixReference" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="fontRef"   type="CT_FontReference"        minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>
