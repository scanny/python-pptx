
``CT_TextParagraphProperties``
==============================


XML Schema excerpt
------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_TextParagraph">
    <xsd:sequence>
      <xsd:element name="pPr"        type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:group   ref="EG_TextRun"  minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="endParaRPr" type="CT_TextCharacterProperties" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextParagraphProperties">
    <xsd:sequence>
      <xsd:element name="lnSpc"  type="CT_TextSpacing"             minOccurs="0"/>
      <xsd:element name="spcBef" type="CT_TextSpacing"             minOccurs="0"/>
      <xsd:element name="spcAft" type="CT_TextSpacing"             minOccurs="0"/>
      <xsd:group   ref="EG_TextBulletColor"                        minOccurs="0"/>
      <xsd:group   ref="EG_TextBulletSize"                         minOccurs="0"/>
      <xsd:group   ref="EG_TextBulletTypeface"                     minOccurs="0"/>
      <xsd:group   ref="EG_TextBullet"                             minOccurs="0"/>
      <xsd:element name="tabLst" type="CT_TextTabStopList"         minOccurs="0"/>
      <xsd:element name="defRPr" type="CT_TextCharacterProperties" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList"  minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="marL"         type="ST_TextMargin"/>
    <xsd:attribute name="marR"         type="ST_TextMargin"/>
    <xsd:attribute name="lvl"          type="ST_TextIndentLevelType"/>
    <xsd:attribute name="indent"       type="ST_TextIndent"/>
    <xsd:attribute name="algn"         type="ST_TextAlignType"/>
    <xsd:attribute name="defTabSz"     type="ST_Coordinate32"/>
    <xsd:attribute name="rtl"          type="xsd:boolean"/>
    <xsd:attribute name="eaLnBrk"      type="xsd:boolean"/>
    <xsd:attribute name="fontAlgn"     type="ST_TextFontAlignType"/>
    <xsd:attribute name="latinLnBrk"   type="xsd:boolean"/>
    <xsd:attribute name="hangingPunct" type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_TextAlignType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="l"/>
      <xsd:enumeration value="ctr"/>
      <xsd:enumeration value="r"/>
      <xsd:enumeration value="just"/>
      <xsd:enumeration value="justLow"/>
      <xsd:enumeration value="dist"/>
      <xsd:enumeration value="thaiDist"/>
    </xsd:restriction>
  </xsd:simpleType>
