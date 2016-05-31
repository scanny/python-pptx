
Font typeface
=============

Overview
--------

PowerPoint allows the font of text elements to be changed from, for example,
Verdana to Arial. This aspect of a font is its *typeface*, as opposed to its
size (e.g. 18 points) or style (e.g. bold, italic).


Minimum viable feature
----------------------

::

    >>> assert isinstance(font, pptx.text.Font)
    >>> font.name
    None
    >>> font.name = 'Verdana'
    >>> font.name
    'Verdana'


Related MS API
--------------

* Font.Name
* Font.NameAscii
* Font.NameComplexScript
* Font.NameFarEast
* Font.NameOther


Protocol
--------

::

    >>> assert isinstance(font, pptx.text.Font)
    >>> font.name
    None
    >>> font.name = 'Verdana'
    >>> font.name
    'Verdana'


XML specimens
-------------

.. highlight:: xml

Here is a representative sample of textbox XML showing the effect of applying
a non-default typeface.

*Baseline default textbox*::

    <p:txBody>
      <a:bodyPr wrap="none" rtlCol="0">
        <a:spAutoFit/>
      </a:bodyPr>
      <a:lstStyle/>
      <a:p>
        <a:r>
          <a:rPr lang="en-US" dirty="0" smtClean="0"/>
          <a:t>Baseline default textbox, no adjustments of any kind</a:t>
        </a:r>
        <a:endParaRPr lang="en-US" dirty="0"/>
      </a:p>
    </p:txBody>

*textbox with typeface applied at shape level*::

    <p:txBody>
      <a:bodyPr wrap="none" rtlCol="0">
        <a:spAutoFit/>
      </a:bodyPr>
      <a:lstStyle/>
      <a:p>
        <a:r>
          <a:rPr lang="en-US" dirty="0" smtClean="0">
            <a:latin typeface="Verdana"/>
            <a:cs typeface="Verdana"/>
          </a:rPr>
          <a:t>A textbox with Verdana typeface applied to shape, not text selection</a:t>
        </a:r>
        <a:endParaRPr lang="en-US" dirty="0">
          <a:latin typeface="Verdana"/>
          <a:cs typeface="Verdana"/>
        </a:endParaRPr>
      </a:p>
    </p:txBody>

*textbox with multiple runs, typeface applied at shape level*::

    <p:txBody>
      <a:bodyPr wrap="none" rtlCol="0">
        <a:spAutoFit/>
      </a:bodyPr>
      <a:lstStyle/>
      <a:p>
        <a:pPr algn="ctr"/>
        <a:r>
          <a:rPr lang="en-US" dirty="0" smtClean="0">
            <a:latin typeface="Arial Black"/>
            <a:cs typeface="Arial Black"/>
          </a:rPr>
          <a:t>textbox with multiple runs having typeface</a:t>
        </a:r>
        <a:br>
          <a:rPr lang="en-US" dirty="0" smtClean="0">
            <a:latin typeface="Arial Black"/>
            <a:cs typeface="Arial Black"/>
          </a:rPr>
        </a:br>
        <a:r>
          <a:rPr lang="en-US" dirty="0" smtClean="0">
            <a:latin typeface="Arial Black"/>
            <a:cs typeface="Arial Black"/>
          </a:rPr>
          <a:t>customized at shape level</a:t>
        </a:r>
        <a:endParaRPr lang="en-US" dirty="0">
          <a:latin typeface="Arial Black"/>
          <a:cs typeface="Arial Black"/>
        </a:endParaRPr>
      </a:p>
    </p:txBody>

*Asian characters, or possibly Asian font being applied*::

    <p:txBody>
      <a:bodyPr wrap="none" rtlCol="0">
        <a:spAutoFit/>
      </a:bodyPr>
      <a:lstStyle/>
      <a:p>
        <a:pPr algn="ctr"/>
        <a:r>
          <a:rPr lang="en-US" dirty="0" smtClean="0">
            <a:latin typeface="Hiragino Sans GB W3"/>
            <a:ea typeface="Hiragino Sans GB W3"/>
            <a:cs typeface="Hiragino Sans GB W3"/>
          </a:rPr>
          <a:t>暒龢加咊晴弗</a:t>
        </a:r>
        <a:endParaRPr lang="en-US" dirty="0">
          <a:latin typeface="Hiragino Sans GB W3"/>
          <a:ea typeface="Hiragino Sans GB W3"/>
          <a:cs typeface="Hiragino Sans GB W3"/>
        </a:endParaRPr>
      </a:p>
    </p:txBody>

*then applying Arial from font pull-down*::

    <p:txBody>
      <a:bodyPr wrap="none" rtlCol="0">
        <a:spAutoFit/>
      </a:bodyPr>
      <a:lstStyle/>
      <a:p>
        <a:pPr algn="ctr"/>
        <a:r>
          <a:rPr lang="en-US" dirty="0" smtClean="0">
            <a:latin typeface="Arial"/>
            <a:ea typeface="Hiragino Sans GB W3"/>
            <a:cs typeface="Arial"/>
          </a:rPr>
          <a:t>暒龢加咊晴弗</a:t>
        </a:r>
        <a:endParaRPr lang="en-US" dirty="0">
          <a:latin typeface="Arial"/>
          <a:ea typeface="Hiragino Sans GB W3"/>
          <a:cs typeface="Arial"/>
        </a:endParaRPr>
      </a:p>
    </p:txBody>


Observations
~~~~~~~~~~~~

* PowerPoint UI always applies typeface customization at run level rather
  than paragraph (defRPr) level, even when applied to shape rather than
  a specific selection of text.
* PowerPoint applies the same typeface to the ``<a:latin>`` and ``<a:cs>``
  tag when a typeface is selected from the font pull-down in the UI.


Related Schema Definitions
--------------------------

.. highlight:: xml

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

  <xsd:complexType name="CT_TextFont">
    <xsd:attribute name="typeface"    type="ST_TextTypeface" use="required"/>
    <xsd:attribute name="panose"      type="s:ST_Panose"     use="optional"/>
    <xsd:attribute name="pitchFamily" type="ST_PitchFamily"  use="optional" default="0"/>
    <xsd:attribute name="charset"     type="xsd:byte"        use="optional" default="1"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_TextTypeface">
    <xsd:restriction base="xsd:string"/>
  </xsd:simpleType>
