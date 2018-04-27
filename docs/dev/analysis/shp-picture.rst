
Picture
=======

PowerPoint allows an image to be added to a slide as a |Picture| shape.


Protocol
--------

Add a picture::

  >>> shapes = Presentation(...).slides[0].shapes
  >>> picture = shapes.add_picture('python.jpg', Inches(1), Inches(1))

Interrogate and set cropping::

  >>> picture.crop_right
  None
  >>> picture.crop_right = 25000  # ---in 1000ths of a percent
  >>> picture.crop_right
  25000.0


XML Specimens
-------------

.. highlight:: xml

Picture shape as added by PowerPoint Mac 2011::

  <p:pic>
    <p:nvPicPr>
      <p:cNvPr id="2" name="Picture 1" descr="sonic.gif"/>
      <p:cNvPicPr>
        <a:picLocks noChangeAspect="1"/>
      </p:cNvPicPr>
      <p:nvPr/>
    </p:nvPicPr>
    <p:blipFill>
      <a:blip r:embed="rId2">
        <a:extLst>
          <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
            <a14:useLocalDpi
              xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
              val="0"/>
          </a:ext>
        </a:extLst>
      </a:blip>
      <a:stretch>
        <a:fillRect/>
      </a:stretch>
    </p:blipFill>
    <p:spPr>
      <a:xfrm>
        <a:off x="2730500" y="1143000"/>
        <a:ext cx="3683000" cy="4572000"/>
      </a:xfrm>
      <a:prstGeom prst="rect">
        <a:avLst/>
      </a:prstGeom>
    </p:spPr>
  </p:pic>


Cropped `pic` (`p:blipFill` child only)::

  <p:blipFill rotWithShape="1">
    <a:blip r:embed="rId2"/>
    <a:srcRect l="9330" t="15904" r="21873" b="28227"/>
    <a:stretch/>
  </p:blipFill>


Schema excerpt
--------------

.. highlight:: xml

::

  <xsd:complexType name="CT_Picture">
    <xsd:sequence>
      <xsd:element name="nvPicPr"  type="CT_PictureNonVisual"/>
      <xsd:element name="blipFill" type="a:CT_BlipFillProperties"/>
      <xsd:element name="spPr"     type="a:CT_ShapeProperties"/>
      <xsd:element name="style"    type="a:CT_ShapeStyle"        minOccurs="0"/>
      <xsd:element name="extLst"   type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_PictureNonVisual">
    <xsd:sequence>
      <xsd:element name="cNvPr"    type="a:CT_NonVisualDrawingProps"/>
      <xsd:element name="cNvPicPr" type="a:CT_NonVisualPictureProperties"/>
      <xsd:element name="nvPr"     type="CT_ApplicationNonVisualDrawingProps"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_BlipFillProperties">
    <xsd:sequence>
      <xsd:element name="blip"    type="CT_Blip"         minOccurs="0"/>
      <xsd:element name="srcRect" type="CT_RelativeRect" minOccurs="0"/>
      <xsd:group   ref="EG_FillModeProperties"           minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="dpi"          type="xsd:unsignedInt"/>
    <xsd:attribute name="rotWithShape" type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"    type="CT_Transform2D"            minOccurs="0"/>
      <xsd:group ref="EG_Geometry"                                 minOccurs="0"/>
      <xsd:group ref="EG_FillProperties"                           minOccurs="0"/>
      <xsd:element name="ln"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:group ref="EG_EffectProperties"                         minOccurs="0"/>
      <xsd:element name="scene3d" type="CT_Scene3D"                minOccurs="0"/>
      <xsd:element name="sp3d"    type="CT_Shape3D"                minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeStyle">
    <xsd:sequence>
      <xsd:element name="lnRef"     type="CT_StyleMatrixReference"/>
      <xsd:element name="fillRef"   type="CT_StyleMatrixReference"/>
      <xsd:element name="effectRef" type="CT_StyleMatrixReference"/>
      <xsd:element name="fontRef"   type="CT_FontReference"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ApplicationNonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="ph"          type="CT_Placeholder"      minOccurs="0"/>
      <xsd:group   ref="a:EG_Media"                              minOccurs="0"/>
      <xsd:element name="custDataLst" type="CT_CustomerDataList" minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="isPhoto"   type="xsd:boolean" default="false"/>
    <xsd:attribute name="userDrawn" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Blip">
    <xsd:sequence>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="alphaBiLevel" type="CT_AlphaBiLevelEffect"/>
        <xsd:element name="alphaCeiling" type="CT_AlphaCeilingEffect"/>
        <xsd:element name="alphaFloor"   type="CT_AlphaFloorEffect"/>
        <xsd:element name="alphaInv"     type="CT_AlphaInverseEffect"/>
        <xsd:element name="alphaMod"     type="CT_AlphaModulateEffect"/>
        <xsd:element name="alphaModFix"  type="CT_AlphaModulateFixedEffect"/>
        <xsd:element name="alphaRepl"    type="CT_AlphaReplaceEffect"/>
        <xsd:element name="biLevel"      type="CT_BiLevelEffect"/>
        <xsd:element name="blur"         type="CT_BlurEffect"/>
        <xsd:element name="clrChange"    type="CT_ColorChangeEffect"/>
        <xsd:element name="clrRepl"      type="CT_ColorReplaceEffect"/>
        <xsd:element name="duotone"      type="CT_DuotoneEffect"/>
        <xsd:element name="fillOverlay"  type="CT_FillOverlayEffect"/>
        <xsd:element name="grayscl"      type="CT_GrayscaleEffect"/>
        <xsd:element name="hsl"          type="CT_HSLEffect"/>
        <xsd:element name="lum"          type="CT_LuminanceEffect"/>
        <xsd:element name="tint"         type="CT_TintEffect"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attributeGroup ref="AG_Blob"/>
    <xsd:attribute name="cstate" type="ST_BlipCompression" default="none"/>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="hlinkClick" type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="hlinkHover" type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="id"     type="ST_DrawingElementId" use="required"/>
    <xsd:attribute name="name"   type="xsd:string"          use="required"/>
    <xsd:attribute name="descr"  type="xsd:string"  default=""/>
    <xsd:attribute name="hidden" type="xsd:boolean" default="false"/>
    <xsd:attribute name="title"  type="xsd:string"  default=""/>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualPictureProperties">
    <xsd:sequence>
      <xsd:element name="picLocks" type="CT_PictureLocking"         minOccurs="0"/>
      <xsd:element name="extLst"   type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="preferRelativeResize" type="xsd:boolean" default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Point2D">
    <xsd:attribute name="x" type="ST_Coordinate" use="required"/>
    <xsd:attribute name="y" type="ST_Coordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PositiveSize2D">
    <xsd:attribute name="cx" type="ST_PositiveCoordinate" use="required"/>
    <xsd:attribute name="cy" type="ST_PositiveCoordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_RelativeRect">
    <xsd:attribute name="l" type="ST_Percentage" default="0%"/>
    <xsd:attribute name="t" type="ST_Percentage" default="0%"/>
    <xsd:attribute name="r" type="ST_Percentage" default="0%"/>
    <xsd:attribute name="b" type="ST_Percentage" default="0%"/>
  </xsd:complexType>

  <xsd:complexType name="CT_StretchInfoProperties">
    <xsd:sequence>
      <xsd:element name="fillRect" type="CT_RelativeRect" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Transform2D">
    <xsd:sequence>
      <xsd:element name="off" type="CT_Point2D" minOccurs="0"/>
      <xsd:element name="ext" type="CT_PositiveSize2D" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="rot" type="ST_Angle" default="0"/>
    <xsd:attribute name="flipH" type="xsd:boolean" default="false"/>
    <xsd:attribute name="flipV" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:attributeGroup name="AG_Blob">
    <xsd:attribute ref="r:embed" default=""/>
    <xsd:attribute ref="r:link" default=""/>
  </xsd:attributeGroup>

  <xsd:group name="EG_EffectProperties">
    <xsd:choice>
      <xsd:element name="effectLst" type="CT_EffectList"/>
      <xsd:element name="effectDag" type="CT_EffectContainer"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_FillModeProperties">
    <xsd:choice>
      <xsd:element name="tile"    type="CT_TileInfoProperties"/>
      <xsd:element name="stretch" type="CT_StretchInfoProperties"/>
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

  <xsd:group name="EG_Geometry">
    <xsd:choice>
      <xsd:element name="custGeom" type="CT_CustomGeometry2D"/>
      <xsd:element name="prstGeom" type="CT_PresetGeometry2D"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_Media">
    <xsd:choice>
      <xsd:element name="audioCd"       type="CT_AudioCD"/>
      <xsd:element name="wavAudioFile"  type="CT_EmbeddedWAVAudioFile"/>
      <xsd:element name="audioFile"     type="CT_AudioFile"/>
      <xsd:element name="videoFile"     type="CT_VideoFile"/>
      <xsd:element name="quickTimeFile" type="CT_QuickTimeFile"/>
    </xsd:choice>
  </xsd:group>
