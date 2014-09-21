
Font - Underline
================

Text in PowerPoint shapes can be formatted with a rich choice of underlining
styles. Two choices control the underline appearance of a font: the underline
style and the underline color. There are 18 available underline styles,
including such choices as single line, wavy, and dot-dash.


Candidate Protocol
------------------

::

    >>> font.underline
    False
    >>> font.underline = True
    >>> font.underline
    SINGLE_LINE (2)
    >>> font.underline = MSO_UNDERLINE.WAVY_DOUBLE_LINE
    >>> font.underline
    WAVY_DOUBLE_LINE (17)
    >>> font.underline = False
    >>> font.underline
    False


XML specimens
-------------

.. highlight:: xml

Run with MSO_UNDERLINE.DOTTED_HEAVY_LINE::

  <a:p>
    <a:r>
      <a:rPr lang="en-US" u="dottedHeavy" dirty="0" smtClean="0"/>
      <a:t>foobar</a:t>
    </a:r>
    <a:endParaRPr lang="en-US" dirty="0"/>
  </a:p>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_TextCharacterProperties">
    <xsd:sequence>
      <xsd:element name="ln"             type="CT_LineProperties" minOccurs="0"/>
      <xsd:group   ref="EG_FillProperties"                        minOccurs="0"/>
      <xsd:group   ref="EG_EffectProperties"                      minOccurs="0"/>
      <xsd:element name="highlight"      type="CT_Color"          minOccurs="0"/>
      <xsd:group   ref="EG_TextUnderlineLine"                     minOccurs="0"/>
      <xsd:group   ref="EG_TextUnderlineFill"                     minOccurs="0"/>
      <xsd:element name="latin"          type="CT_TextFont"       minOccurs="0"/>
      <xsd:element name="ea"             type="CT_TextFont"       minOccurs="0"/>
      <xsd:element name="cs"             type="CT_TextFont"       minOccurs="0"/>
      <xsd:element name="sym"            type="CT_TextFont"       minOccurs="0"/>
      <xsd:element name="hlinkClick"     type="CT_Hyperlink"      minOccurs="0"/>
      <xsd:element name="hlinkMouseOver" type="CT_Hyperlink"      minOccurs="0"/>
      <xsd:element name="rtl"            type="CT_Boolean"        minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_OfficeArtExtensionList" minOccurs="0" />
    </xsd:sequence>
    <xsd:attribute name="kumimoji"   type="xsd:boolean"/>
    <xsd:attribute name="lang"       type="s:ST_Lang"/>
    <xsd:attribute name="altLang"    type="s:ST_Lang"/>
    <xsd:attribute name="sz"         type="ST_TextFontSize"/>
    <xsd:attribute name="b"          type="xsd:boolean"/>
    <xsd:attribute name="i"          type="xsd:boolean"/>
    <xsd:attribute name="u"          type="ST_TextUnderlineType"/>
    <xsd:attribute name="strike"     type="ST_TextStrikeType"/>
    <xsd:attribute name="kern"       type="ST_TextNonNegativePoint"/>
    <xsd:attribute name="cap"        type="ST_TextCapsType"/>
    <xsd:attribute name="spc"        type="ST_TextPoint"/>
    <xsd:attribute name="normalizeH" type="xsd:boolean"/>
    <xsd:attribute name="baseline"   type="ST_Percentage"/>
    <xsd:attribute name="noProof"    type="xsd:boolean"/>
    <xsd:attribute name="dirty"      type="xsd:boolean"     default="true"/>
    <xsd:attribute name="err"        type="xsd:boolean"     default="false"/>
    <xsd:attribute name="smtClean"   type="xsd:boolean"     default="true"/>
    <xsd:attribute name="smtId"      type="xsd:unsignedInt" default="0"/>
    <xsd:attribute name="bmk"        type="xsd:string"/>
  </xsd:complexType>

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

  <xsd:group name="EG_TextUnderlineLine">
    <xsd:choice>
      <xsd:element name="uLnTx" type="CT_TextUnderlineLineFollowText"/>
      <xsd:element name="uLn"   type="CT_LineProperties" minOccurs="0"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_TextUnderlineFill">
    <xsd:choice>
      <xsd:element name="uFillTx" type="CT_TextUnderlineFillFollowText"/>
      <xsd:element name="uFill"   type="CT_TextUnderlineFillGroupWrapper"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_TextUnderlineLineFollowText"/>

  <xsd:complexType name="CT_TextUnderlineFillFollowText"/>

  <xsd:complexType name="CT_TextUnderlineFillGroupWrapper">
    <xsd:group ref="EG_FillProperties"/>
  </xsd:complexType>
