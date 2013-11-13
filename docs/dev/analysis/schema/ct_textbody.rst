``CT_TextBody``
===============

.. highlight:: xml

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 15, 50

   Spec Name    , Shape Text Body
   Tag(s)       , p:txBody
   Namespace    , presentationml (pml.xsd)
   Spec Section , 19.3.1.51


attributes
----------

None.


child elements
--------------

=========  ====  ================================
name        #    type
=========  ====  ================================
bodyPr      1    :doc:`ct_textbodyproperties`
lstStyle    ?    CT_TextListStyle
p           \+   :doc:`ct_textparagraph`
=========  ====  ================================


Spec text
---------

   This element specifies the existence of text to be contained within the
   corresponding shape. All visible text and visible text related properties
   are contained within this element. There can be multiple paragraphs and
   within paragraphs multiple runs of text.


Schema excerpt
--------------

::

  <xsd:complexType name="CT_TextBody">
    <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="p"        type="CT_TextParagraph"      minOccurs="1" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

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

  <xsd:complexType name="CT_TextListStyle">
    <xsd:sequence>
      <xsd:element name="defPPr"  type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="lvl1pPr" type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="lvl2pPr" type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="lvl3pPr" type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="lvl4pPr" type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="lvl5pPr" type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="lvl6pPr" type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="lvl7pPr" type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="lvl8pPr" type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="lvl9pPr" type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList"  minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextParagraph">
    <xsd:sequence>
      <xsd:element name="pPr"        type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_TextRun"  minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="endParaRPr" type="CT_TextCharacterProperties" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>
