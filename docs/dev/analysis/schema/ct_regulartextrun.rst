
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


attributes
----------

None.


child elements
--------------

=====  ===  ==============================================================
name    #   type
=====  ===  ==============================================================
rPr     ?   :doc:`ct_textcharacterproperties`
t       1   xsd:string
=====  ===  ==============================================================


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
      <xsd:element name="rPr" type="CT_TextCharacterProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="t" type="xsd:string" minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextCharacterProperties">
    <xsd:sequence>
      <xsd:element name="ln"             type="CT_LineProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:group    ref="EG_FillProperties"                       minOccurs="0" maxOccurs="1"/>
      <xsd:group    ref="EG_EffectProperties"                     minOccurs="0" maxOccurs="1"/>
      <xsd:element name="highlight"      type="CT_Color"          minOccurs="0" maxOccurs="1"/>
      <xsd:group    ref="EG_TextUnderlineLine"                    minOccurs="0" maxOccurs="1"/>
      <xsd:group    ref="EG_TextUnderlineFill"                    minOccurs="0" maxOccurs="1"/>
      <xsd:element name="latin"          type="CT_TextFont"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="ea"             type="CT_TextFont"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="cs"             type="CT_TextFont"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sym"            type="CT_TextFont"       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hlinkClick"     type="CT_Hyperlink"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hlinkMouseOver" type="CT_Hyperlink"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="rtl"            type="CT_Boolean"        minOccurs="0"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
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
