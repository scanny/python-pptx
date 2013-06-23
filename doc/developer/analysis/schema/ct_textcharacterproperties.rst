##############################
``CT_TextCharacterProperties``
##############################

.. highlight:: xml

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 25, 50

   Schema Name  , CT_TextCharacterProperties
   Spec Name    , Text Run Properties
   Tag(s)       , a:rPr
   Namespace    , drawingml (dml-main.xsd)
   Schema Line  , 2907
   Spec Section , 21.1.2.3.9 rPr (Text Run Properties)


Resources
=========

* ISO-IEC-29500-1, Section 21.1.2.3.9 rPr (Text Run Properties)


Spec text
=========

   This element contains all run level text properties for the text runs within
   a containing paragraph.


Schema excerpt
==============

::

  <xsd:complexType name="CT_TextCharacterProperties">
    <xsd:sequence>
      <xsd:element name="ln"        type="CT_LineProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_FillProperties"                   minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_EffectProperties"                 minOccurs="0" maxOccurs="1"/>
      <xsd:element name="highlight"      type="CT_Color"     minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_TextUnderlineLine"                minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_TextUnderlineFill"                minOccurs="0" maxOccurs="1"/>
      <xsd:element name="latin"          type="CT_TextFont"  minOccurs="0" maxOccurs="1"/>
      <xsd:element name="ea"             type="CT_TextFont"  minOccurs="0" maxOccurs="1"/>
      <xsd:element name="cs"             type="CT_TextFont"  minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sym"            type="CT_TextFont"  minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hlinkClick"     type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hlinkMouseOver" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="rtl"            type="CT_Boolean"   minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_OfficeArtExtensionList" minOccurs="0"  maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="kumimoji"   type="xsd:boolean"             use="optional"/>
    <xsd:attribute name="lang"       type="s:ST_Lang"               use="optional"/>
    <xsd:attribute name="altLang"    type="s:ST_Lang"               use="optional"/>
    <xsd:attribute name="sz"         type="ST_TextFontSize"         use="optional"/>
    <xsd:attribute name="b"          type="xsd:boolean"             use="optional"/>
    <xsd:attribute name="i"          type="xsd:boolean"             use="optional"/>
    <xsd:attribute name="u"          type="ST_TextUnderlineType"    use="optional"/>
    <xsd:attribute name="strike"     type="ST_TextStrikeType"       use="optional"/>
    <xsd:attribute name="kern"       type="ST_TextNonNegativePoint" use="optional"/>
    <xsd:attribute name="cap"        type="ST_TextCapsType"         use="optional"/>
    <xsd:attribute name="spc"        type="ST_TextPoint"            use="optional"/>
    <xsd:attribute name="normalizeH" type="xsd:boolean"             use="optional"/>
    <xsd:attribute name="baseline"   type="ST_Percentage"           use="optional"/>
    <xsd:attribute name="noProof"    type="xsd:boolean"             use="optional"/>
    <xsd:attribute name="dirty"      type="xsd:boolean"             use="optional" default="true"/>
    <xsd:attribute name="err"        type="xsd:boolean"             use="optional" default="false"/>
    <xsd:attribute name="smtClean"   type="xsd:boolean"             use="optional" default="true"/>
    <xsd:attribute name="smtId"      type="xsd:unsignedInt"         use="optional" default="0"/>
    <xsd:attribute name="bmk"        type="xsd:string"              use="optional"/>
  </xsd:complexType>

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

  <xsd:group name="EG_TextUnderlineLine">
    <xsd:choice>
      <xsd:element name="uLnTx" type="CT_TextUnderlineLineFollowText"/>
      <xsd:element name="uLn"   type="CT_LineProperties" minOccurs="0" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_TextUnderlineFill">
    <xsd:choice>
      <xsd:element name="uFillTx" type="CT_TextUnderlineFillFollowText"/>
      <xsd:element name="uFill"   type="CT_TextUnderlineFillGroupWrapper"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_TextFont">
    <xsd:attribute name="typeface"    type="ST_TextTypeface" use="required"/>
    <xsd:attribute name="panose"      type="s:ST_Panose"     use="optional"/>
    <xsd:attribute name="pitchFamily" type="ST_PitchFamily"  use="optional" default="0"/>
    <xsd:attribute name="charset"     type="xsd:byte"        use="optional" default="1"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Hyperlink">
    <xsd:sequence>
      <xsd:element name="snd"    type="CT_EmbeddedWAVAudioFile"   minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute ref="r:id" use="optional"/>
      <xsd:attribute name="invalidUrl"     type="xsd:string"  use="optional" default=""/>
      <xsd:attribute name="action"         type="xsd:string"  use="optional" default=""/>
      <xsd:attribute name="tgtFrame"       type="xsd:string"  use="optional" default=""/>
      <xsd:attribute name="tooltip"        type="xsd:string"  use="optional" default=""/>
      <xsd:attribute name="history"        type="xsd:boolean" use="optional" default="true"/>
      <xsd:attribute name="highlightClick" type="xsd:boolean" use="optional" default="false"/>
      <xsd:attribute name="endSnd"         type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Panose">
    <xsd:restriction base="xsd:hexBinary">
      <xsd:length value="10"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Percentage">
    <xsd:union memberTypes="ST_PercentageDecimal s:ST_Percentage"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextCapsType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="small"/>
      <xsd:enumeration value="all"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextNonNegativePoint">
    <xsd:restriction base="xsd:int">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="400000"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextPoint">
    <xsd:union memberTypes="ST_TextPointUnqualified s:ST_UniversalMeasure"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextStrikeType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="noStrike"/>
      <xsd:enumeration value="sngStrike"/>
      <xsd:enumeration value="dblStrike"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextUnderlineType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="words"/>
      <xsd:enumeration value="sng"/>
      <xsd:enumeration value="dbl"/>
      <xsd:enumeration value="heavy"/>
      <xsd:enumeration value="dotted"/>
      <xsd:enumeration value="dottedHeavy"/>
      <xsd:enumeration value="dash"/>
      <xsd:enumeration value="dashHeavy"/>
      <xsd:enumeration value="dashLong"/>
      <xsd:enumeration value="dashLongHeavy"/>
      <xsd:enumeration value="dotDash"/>
      <xsd:enumeration value="dotDashHeavy"/>
      <xsd:enumeration value="dotDotDash"/>
      <xsd:enumeration value="dotDotDashHeavy"/>
      <xsd:enumeration value="wavy"/>
      <xsd:enumeration value="wavyHeavy"/>
      <xsd:enumeration value="wavyDbl"/>
    </xsd:restriction>
  </xsd:simpleType>
