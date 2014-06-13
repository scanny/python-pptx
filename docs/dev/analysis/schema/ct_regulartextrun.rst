
``CT_RegularTextRun``
=====================

.. highlight:: xml

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 15, 50

   Spec Name    , Text Run
   Tag(s)       , a:r
   Namespace    , drawingml (dml-main.xsd)
   Spec Section , 21.1.2.3.8


Spec text
---------

    This element specifies the presence of a run of text within the containing
    text body. The run element is the lowest level text separation mechanism
    within a text body. A text run can contain text run properties associated
    with the run. If no properties are listed then properties specified in the
    defRPr element are used.


Schema excerpt
--------------

::

  <xsd:complexType name="CT_RegularTextRun">
    <xsd:sequence>
      <xsd:element name="rPr" type="CT_TextCharacterProperties" minOccurs="0"/>
      <xsd:element name="t"   type="xsd:string"/>
    </xsd:sequence>
  </xsd:complexType>

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
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
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
