====================
``CT_TextParagraph``
====================

.. highlight:: xml

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 15, 50

   Schema Name  , CT_TextParagraph
   Spec Name    , Text Paragraph
   Tag(s)       , a:p
   Namespace    , drawingml (dml-main.xsd)
   Schema Line  , 2531 dml
   Spec Section , 21.1.2.2.6


Analysis
========

XPath expression from `<p:txBody>` is ``./a:p``

`<p:txBody>` must always contain at least one `<a:p>` element.


attributes
^^^^^^^^^^

None.


child elements
^^^^^^^^^^^^^^

==========  ===  ==========================  ==========
name         #    type                        line
==========  ===  ==========================  ==========
pPr          1   CT_TextParagraphProperties  2963 dml
r           \*   EG_TextRun                  2997 dml
\| br
\| fld
endParaRPr   ?   CT_TextCharacterProperties  2842 dml
==========  ===  ==========================  ==========


Spec text
^^^^^^^^^

   This element specifies the presence of a paragraph of text within the
   containing text body. The paragraph is the highest level text separation
   mechanism within a text body. A paragraph can contain text paragraph
   properties associated with the paragraph. If no properties are listed then
   properties specified in the defPPr element are used.


Schema excerpt
^^^^^^^^^^^^^^

::

  <xsd:complexType name="CT_TextParagraph">
    <xsd:sequence>
      <xsd:element name="pPr"        type="CT_TextParagraphProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_TextRun"  minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="endParaRPr" type="CT_TextCharacterProperties" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextParagraphProperties">
    <xsd:sequence>
      <xsd:element name="lnSpc"  type="CT_TextSpacing"             minOccurs="0" maxOccurs="1"/>
      <xsd:element name="spcBef" type="CT_TextSpacing"             minOccurs="0" maxOccurs="1"/>
      <xsd:element name="spcAft" type="CT_TextSpacing"             minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_TextBulletColor"                          minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_TextBulletSize"                           minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_TextBulletTypeface"                       minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_TextBullet"                               minOccurs="0" maxOccurs="1"/>
      <xsd:element name="tabLst" type="CT_TextTabStopList"         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="defRPr" type="CT_TextCharacterProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList"  minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="marL"         type="ST_TextMargin"          use="optional"/>
    <xsd:attribute name="marR"         type="ST_TextMargin"          use="optional"/>
    <xsd:attribute name="lvl"          type="ST_TextIndentLevelType" use="optional"/>
    <xsd:attribute name="indent"       type="ST_TextIndent"          use="optional"/>
    <xsd:attribute name="algn"         type="ST_TextAlignType"       use="optional"/>
    <xsd:attribute name="defTabSz"     type="ST_Coordinate32"        use="optional"/>
    <xsd:attribute name="rtl"          type="xsd:boolean"            use="optional"/>
    <xsd:attribute name="eaLnBrk"      type="xsd:boolean"            use="optional"/>
    <xsd:attribute name="fontAlgn"     type="ST_TextFontAlignType"   use="optional"/>
    <xsd:attribute name="latinLnBrk"   type="xsd:boolean"            use="optional"/>
    <xsd:attribute name="hangingPunct" type="xsd:boolean"            use="optional"/>
  </xsd:complexType>

  <xsd:group name="EG_TextRun">
    <xsd:choice>
      <xsd:element name="r" type="CT_RegularTextRun"/>
      <xsd:element name="br" type="CT_TextLineBreak"/>
      <xsd:element name="fld" type="CT_TextField"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_TextAutofit">
    <xsd:choice>
      <xsd:element name="noAutofit"   type="CT_TextNoAutofit"     />
      <xsd:element name="normAutofit" type="CT_TextNormalAutofit" />
      <xsd:element name="spAutoFit"   type="CT_TextShapeAutofit"  />
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_RegularTextRun">
    <xsd:sequence>
      <xsd:element name="rPr" type="CT_TextCharacterProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="t" type="xsd:string" minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextLineBreak">
    <xsd:sequence>
      <xsd:element name="rPr" type="CT_TextCharacterProperties" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextCharacterProperties">
    <xsd:sequence>
      <xsd:element name="ln" type="CT_LineProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_FillProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_EffectProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="highlight" type="CT_Color" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_TextUnderlineLine" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_TextUnderlineFill" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="latin" type="CT_TextFont" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="ea" type="CT_TextFont" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="cs" type="CT_TextFont" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sym" type="CT_TextFont" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hlinkClick" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hlinkMouseOver" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="rtl" type="CT_Boolean" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="kumimoji" type="xsd:boolean" use="optional"/>
    <xsd:attribute name="lang" type="s:ST_Lang" use="optional"/>
    <xsd:attribute name="altLang" type="s:ST_Lang" use="optional"/>
    <xsd:attribute name="sz" type="ST_TextFontSize" use="optional"/>
    <xsd:attribute name="b" type="xsd:boolean" use="optional"/>
    <xsd:attribute name="i" type="xsd:boolean" use="optional"/>
    <xsd:attribute name="u" type="ST_TextUnderlineType" use="optional"/>
    <xsd:attribute name="strike" type="ST_TextStrikeType" use="optional"/>
    <xsd:attribute name="kern" type="ST_TextNonNegativePoint" use="optional"/>
    <xsd:attribute name="cap" type="ST_TextCapsType" use="optional"/>
    <xsd:attribute name="spc" type="ST_TextPoint" use="optional"/>
    <xsd:attribute name="normalizeH" type="xsd:boolean" use="optional"/>
    <xsd:attribute name="baseline" type="ST_Percentage" use="optional"/>
    <xsd:attribute name="noProof" type="xsd:boolean" use="optional"/>
    <xsd:attribute name="dirty" type="xsd:boolean" use="optional" default="true"/>
    <xsd:attribute name="err" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="smtClean" type="xsd:boolean" use="optional" default="true"/>
    <xsd:attribute name="smtId" type="xsd:unsignedInt" use="optional" default="0"/>
    <xsd:attribute name="bmk" type="xsd:string" use="optional"/>
  </xsd:complexType>


