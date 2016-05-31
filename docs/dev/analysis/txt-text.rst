
Text
====

All text in PowerPoint appears in a shape. Of the six shape types, only
an AutoShape can directly contain text. This includes text placeholders, text
boxes, and geometric auto shapes. These correspond to shapes based on the
``<p:sp>`` XML element.

Each auto shape has a text frame, represented by a |TextFrame| object. A text
frame contains one or more paragraphs, and each paragraph contains a sequence
of zero or more run, line break, or field elements (``<a:r>``, ``<a:br>``,
and ``<a:fld>``, respectively).


Notes
-----

* include a:fld/a:t contents in Paragraph.text


XML specimens
-------------

.. highlight:: xml

Default text box with text, a line break, and a field::

  <p:txBody>
    <a:bodyPr wrap="none" rtlCol="0">
      <a:spAutoFit/>
    </a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" dirty="0" smtClean="0"/>
        <a:t>This text box has a break here -&gt;</a:t>
      </a:r>
      <a:br>
        <a:rPr lang="en-US" dirty="0" smtClean="0"/>
      </a:br>
      <a:r>
        <a:rPr lang="en-US" dirty="0" smtClean="0"/>
        <a:t>and a Field here -&gt; </a:t>
      </a:r>
      <a:fld id="{466D331B-39A1-D247-9EB6-2F41F44AA032}" type="datetime2">
        <a:rPr lang="en-US" smtClean="0"/>
        <a:t>Monday, September 15, 14</a:t>
      </a:fld>
      <a:endParaRPr lang="en-US" dirty="0"/>
    </a:p>
  </p:txBody>


Schema excerpt
--------------

::

  <xsd:complexType name="CT_TextBody">
    <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle"      minOccurs="0"/>
      <xsd:element name="p"        type="CT_TextParagraph"      maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextBodyProperties">
    <xsd:sequence>
      <xsd:element name="prstTxWarp" type="CT_PresetTextShape"        minOccurs="0"/>
      <xsd:group ref="EG_TextAutofit"                                 minOccurs="0"/>
      <xsd:element name="scene3d"    type="CT_Scene3D"                minOccurs="0"/>
      <xsd:group ref="EG_Text3D"                                      minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="rot"              type="ST_Angle"/>
    <xsd:attribute name="spcFirstLastPara" type="xsd:boolean"/>
    <xsd:attribute name="vertOverflow"     type="ST_TextVertOverflowType"/>
    <xsd:attribute name="horzOverflow"     type="ST_TextHorzOverflowType"/>
    <xsd:attribute name="vert"             type="ST_TextVerticalType"/>
    <xsd:attribute name="wrap"             type="ST_TextWrappingType"/>
    <xsd:attribute name="lIns"             type="ST_Coordinate32"/>
    <xsd:attribute name="tIns"             type="ST_Coordinate32"/>
    <xsd:attribute name="rIns"             type="ST_Coordinate32"/>
    <xsd:attribute name="bIns"             type="ST_Coordinate32"/>
    <xsd:attribute name="numCol"           type="ST_TextColumnCount"/>
    <xsd:attribute name="spcCol"           type="ST_PositiveCoordinate32"/>
    <xsd:attribute name="rtlCol"           type="xsd:boolean"/>
    <xsd:attribute name="fromWordArt"      type="xsd:boolean"/>
    <xsd:attribute name="anchor"           type="ST_TextAnchoringType"/>
    <xsd:attribute name="anchorCtr"        type="xsd:boolean"/>
    <xsd:attribute name="forceAA"          type="xsd:boolean"/>
    <xsd:attribute name="upright"          type="xsd:boolean" default="false"/>
    <xsd:attribute name="compatLnSpc"      type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TextListStyle">
    <xsd:sequence>
      <xsd:element name="defPPr"  type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl1pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl2pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl3pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl4pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl5pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl6pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl7pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl8pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="lvl9pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList"  minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextParagraph">
    <xsd:sequence>
      <xsd:element name="pPr"        type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:group    ref="EG_TextRun"                                   minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="endParaRPr" type="CT_TextCharacterProperties" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="EG_TextRun">
    <xsd:choice>
      <xsd:element name="r"   type="CT_RegularTextRun"/>
      <xsd:element name="br"  type="CT_TextLineBreak"/>
      <xsd:element name="fld" type="CT_TextField"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_RegularTextRun">
    <xsd:sequence>
      <xsd:element name="rPr" type="CT_TextCharacterProperties" minOccurs="0"/>
      <xsd:element name="t"   type="xsd:string"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextLineBreak">
    <xsd:sequence>
      <xsd:element name="rPr" type="CT_TextCharacterProperties" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextField">
    <xsd:sequence>
      <xsd:element name="rPr" type="CT_TextCharacterProperties" minOccurs="0"/>
      <xsd:element name="pPr" type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:element name="t"   type="xsd:string"                 minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="id"   type="s:ST_Guid"  use="required"/>
    <xsd:attribute name="type" type="xsd:string"/>
  </xsd:complexType>
