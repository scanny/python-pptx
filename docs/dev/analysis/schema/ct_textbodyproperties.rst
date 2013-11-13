
``CT_TextBodyProperties``
=========================

.. highlight:: xml

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 15, 50

   Spec Name    , Body Properties
   Tag(s)       , a:bodyPr
   Namespace    , drawingml (dml-main.xsd)
   Spec Section , 21.1.2.1.1


Attributes
----------

================  ===  =======================
name               #   type
================  ===  =======================
rot                ?   ST_Angle
spcFirstLastPara   ?   xsd:boolean
vertOverflow       ?   ST_TextVertOverflowType
horzOverflow       ?   ST_TextHorzOverflowType
vert               ?   ST_TextVerticalType
wrap               ?   ST_TextWrappingType
lIns               ?   ST_Coordinate32
tIns               ?   ST_Coordinate32
rIns               ?   ST_Coordinate32
bIns               ?   ST_Coordinate32
numCol             ?   ST_TextColumnCount
spcCol             ?   ST_PositiveCoordinate32
rtlCol             ?   xsd:boolean
fromWordArt        ?   xsd:boolean
anchor             ?   ST_TextAnchoringType
anchorCtr          ?   xsd:boolean
forceAA            ?   xsd:boolean
upright            ?   xsd:boolean
compatLnSpc        ?   xsd:boolean
================  ===  =======================


Child elements
--------------

==============  ===  =========================
name             #   type
==============  ===  =========================
prstTxWarp       ?   CT_PresetTextShape
EG_TextAutoFit   ?
scene3d          ?   CT_Scene3D
EG_Text3D        ?
extLst           ?   CT_OfficeArtExtensionList
==============  ===  =========================


Spec text
---------

    This element defines the body properties for the text body within a shape.


Schema excerpt
--------------

::

  <xsd:complexType name="CT_TextBodyProperties">
    <xsd:sequence>
      <xsd:element name="prstTxWarp" type="CT_PresetTextShape"        minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_TextAutofit"                                 minOccurs="0" maxOccurs="1"/>
      <xsd:element name="scene3d"    type="CT_Scene3D"                minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_Text3D"                                      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"     type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="rot"              type="ST_Angle"                use="optional"/>
    <xsd:attribute name="spcFirstLastPara" type="xsd:boolean"             use="optional"/>
    <xsd:attribute name="vertOverflow"     type="ST_TextVertOverflowType" use="optional"/>
    <xsd:attribute name="horzOverflow"     type="ST_TextHorzOverflowType" use="optional"/>
    <xsd:attribute name="vert"             type="ST_TextVerticalType"     use="optional"/>
    <xsd:attribute name="wrap"             type="ST_TextWrappingType"     use="optional"/>
    <xsd:attribute name="lIns"             type="ST_Coordinate32"         use="optional"/>
    <xsd:attribute name="tIns"             type="ST_Coordinate32"         use="optional"/>
    <xsd:attribute name="rIns"             type="ST_Coordinate32"         use="optional"/>
    <xsd:attribute name="bIns"             type="ST_Coordinate32"         use="optional"/>
    <xsd:attribute name="numCol"           type="ST_TextColumnCount"      use="optional"/>
    <xsd:attribute name="spcCol"           type="ST_PositiveCoordinate32" use="optional"/>
    <xsd:attribute name="rtlCol"           type="xsd:boolean"             use="optional"/>
    <xsd:attribute name="fromWordArt"      type="xsd:boolean"             use="optional"/>
    <xsd:attribute name="anchor"           type="ST_TextAnchoringType"    use="optional"/>
    <xsd:attribute name="anchorCtr"        type="xsd:boolean"             use="optional"/>
    <xsd:attribute name="forceAA"          type="xsd:boolean"             use="optional"/>
    <xsd:attribute name="upright"          type="xsd:boolean"             use="optional" default="false"/>
    <xsd:attribute name="compatLnSpc"      type="xsd:boolean"             use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PresetTextShape">
    <xsd:sequence>
      <xsd:element name="avLst" type="CT_GeomGuideList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="prst" type="ST_TextShapeType" use="required"/>
  </xsd:complexType>

  <xsd:group name="EG_TextAutofit">
    <xsd:choice>
      <xsd:element name="noAutofit"   type="CT_TextNoAutofit"/>
      <xsd:element name="normAutofit" type="CT_TextNormalAutofit"/>
      <xsd:element name="spAutoFit"   type="CT_TextShapeAutofit"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_TextNormalAutofit">
    <xsd:attribute name="fontScale"      type="ST_TextFontScalePercentOrPercentString" use="optional" default="100%"/>
    <xsd:attribute name="lnSpcReduction" type="ST_TextSpacingPercentOrPercentString"   use="optional" default="0%"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TextShapeAutofit"/>

  <xsd:complexType name="CT_TextNoAutofit"/>

  <xsd:complexType name="CT_Scene3D">
    <xsd:sequence>
      <xsd:element name="camera"   type="CT_Camera"                 minOccurs="1" maxOccurs="1"/>
      <xsd:element name="lightRig" type="CT_LightRig"               minOccurs="1" maxOccurs="1"/>
      <xsd:element name="backdrop" type="CT_Backdrop"               minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"   type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="EG_Text3D">
    <xsd:choice>
      <xsd:element name="sp3d"   type="CT_Shape3D"  minOccurs="1" maxOccurs="1"/>
      <xsd:element name="flatTx" type="CT_FlatText" minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_Shape3D">
    <xsd:sequence>
      <xsd:element name="bevelT"       type="CT_Bevel" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="bevelB"       type="CT_Bevel" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extrusionClr" type="CT_Color" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="contourClr"   type="CT_Color" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"       type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="z"            type="ST_Coordinate"         use="optional" default="0"/>
    <xsd:attribute name="extrusionH"   type="ST_PositiveCoordinate" use="optional" default="0"/>
    <xsd:attribute name="contourW"     type="ST_PositiveCoordinate" use="optional" default="0"/>
    <xsd:attribute name="prstMaterial" type="ST_PresetMaterialType" use="optional" default="warmMatte"/>
  </xsd:complexType>

  <xsd:complexType name="CT_FlatText">
    <xsd:attribute name="z" type="ST_Coordinate" use="optional" default="0"/>
  </xsd:complexType>

  <xsd:complexType name="CT_OfficeArtExtensionList">
    <xsd:sequence>
      <xsd:group ref="EG_OfficeArtExtensionList" minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:simpleType name="ST_Angle">
    <xsd:restriction base="xsd:int"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Coordinate32">
    <xsd:union memberTypes="ST_Coordinate32Unqualified s:ST_UniversalMeasure"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Coordinate32Unqualified">
    <xsd:restriction base="xsd:int"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_PositiveCoordinate32">
    <xsd:restriction base="ST_Coordinate32Unqualified">
      <xsd:minInclusive value="0"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextAnchoringType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="t"/>
      <xsd:enumeration value="ctr"/>
      <xsd:enumeration value="b"/>
      <xsd:enumeration value="just"/>
      <xsd:enumeration value="dist"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextColumnCount">
    <xsd:restriction base="xsd:int">
      <xsd:minInclusive value="1"/>
      <xsd:maxInclusive value="16"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextHorzOverflowType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="overflow"/>
      <xsd:enumeration value="clip"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextVertOverflowType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="overflow"/>
      <xsd:enumeration value="ellipsis"/>
      <xsd:enumeration value="clip"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextVerticalType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="horz"/>
      <xsd:enumeration value="vert"/>
      <xsd:enumeration value="vert270"/>
      <xsd:enumeration value="wordArtVert"/>
      <xsd:enumeration value="eaVert"/>
      <xsd:enumeration value="mongolianVert"/>
      <xsd:enumeration value="wordArtVertRtl"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextWrappingType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="square"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_UniversalMeasure">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)"/>
    </xsd:restriction>
  </xsd:simpleType>
