
Text - Fit text to shape
========================

An AutoShape has a text frame, referred to in the PowerPoint UI as the
shape's *Text Box*. One of the settings provided is *Autofit*, which can be
one of "Do not autofit", "Resize text to fit shape", or "Resize shape to fit
text". The scope of this analysis is how best to provide an alternative to
the "Resize text to fit shape" behavior that simply resizes the text to the
largest point size that will fit entirely within the shape extents.

This produces a similar visual effect, but the "auto-size" behavior does not
persist to later changes to the text or shape. It is just as though the user
had reduced the font size just until all the text fit within the shape.


Candidate Protocol
------------------

Shape size and text are set before calling :meth:`.TextFrame.fit_text`::

    >>> shape.width, shape.height = cx, cy
    >>> shape.text = 'Lorem ipsum .. post facto.'
    >>> text_frame = shape.text_frame
    >>> text_frame.auto_size, text_frame.word_wrap
    (None, False)

Calling :meth:`TextFrame.fit_text` sets auto-size to MSO_AUTO_SIZE.NONE,
turns word-wrap on, and sets the font size of all text in the shape to the
maximum size that will fit entirely within the shape, not to exceed the
optional *max_size*. The default *max_size* is 18pt. The font size is applied
directly on each run-level element in the shape. The path to the matching
font file must be specified::

    >>> text_frame.fit_text('calibriz.ttf', max_size=18)
    >>> text_frame.auto_size, text_frame.word_wrap
    (MSO_AUTO_SIZE.NONE (0), True)
    >>> text_frame.paragraphs[0].runs[0].font.size.pt
    10


Current constraints
-------------------

:meth:`.TextFrame.fit_text`
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

* Path to font file must be provided manually.
* Bold and italic variant is selected unconditionally.
* Calibri typeface name is selected unconditionally.
* User must manually set the font typeface of all shape text to match the
  provided font file.
* Font typeface and variant is assumed to be uniform.
* Font size is made uniform, without regard to existing differences in font
  size in the shape text.
* Line breaks are replaced with a single space.


Incremental enhancements
------------------------

* **Allow font file to be specified.** This allows
  :meth:`.TextFrame.fit_text` to be run from `pip` installed version.
  OS-specific file locations and names can be determined by client.

* **Allow bold and italic to be specified.** This allows the client to
  determine whether to apply the bold and/or italic typeface variant to the
  text.

* **Allow typeface name to be specified.** This allows the client to
  determine which base typeface is applied to the text.

* **Auto-locate font file based on typeface and variants.** Relax requirement
  for client to specify font file, looking it up in current system based on
  system type, typeface name, and variants.


PowerPoint behavior
-------------------

* PowerPoint shrinks text in whole-number font sizes.

* When assigning a font size to a shape, PowerPoint applies that font size at
  the run level, adding a `sz` attribute to the `<a:rPr>` element for every
  content child of every `<a:p>` element in the shape. The `<a:endParaRPr>`
  element in each paragraph also gets a `sz` attribute set to that size.


XML specimens
-------------

.. highlight:: xml

``<p:txBody>`` for default new textbox::

  <p:txBody>
    <a:bodyPr wrap="none">
      <a:spAutoFit/>  <!-- fit shape to text -->
    </a:bodyPr>
    <a:lstStyle/>
    <a:p/>
  </p:txBody>

8" x 0.5" text box, default margins, defaulting to 18pt "full-size" text,
auto-reduced to 12pt. ``<a:t>`` element text wrapped for compact display::

  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0">
      <a:noAutofit/>
    </a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz=1200 dirty="0" smtClean="0"/>
        <a:t>The art and craft of designing typefaces is called type design.
             Designers of typefaces are called type designers and are often
             employed by type foundries. In digital typography, type
             designers are sometimes also called font developers or font
             designers.</a:t>
      </a:r>
      <a:endParaRPr lang="en-US" sz=1200 dirty="0"/>
    </a:p>
  </p:txBody>


Related Schema Definitions
--------------------------

::

  <xsd:complexType name="CT_TextBody">
    <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle" minOccurs="0"/>
      <xsd:element name="p"        type="CT_TextParagraph" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextBodyProperties">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="prstTxWarp"  type="CT_PresetTextShape"        minOccurs="0"/>
      <xsd:choice minOccurs="0">      <!-- EG_TextAutofit -->
        <xsd:element name="noAutofit"   type="CT_TextNoAutofit"/>
        <xsd:element name="normAutofit" type="CT_TextNormalAutofit"/>
        <xsd:element name="spAutoFit"   type="CT_TextShapeAutofit"/>
      </xsd:choice>
      <xsd:element name="scene3d"     type="CT_Scene3D"                minOccurs="0"/>
      <xsd:choice minOccurs="0">      <!-- EG_Text3D -->
        <xsd:element name="sp3d"        type="CT_Shape3D"/>
        <xsd:element name="flatTx"      type="CT_FlatText"/>
      </xsd:choice>
      <xsd:element name="extLst"      type="CT_OfficeArtExtensionList" minOccurs="0"/>
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

  <xsd:complexType name="CT_TextNoAutofit"/>

  <xsd:complexType name="CT_TextParagraph">
    <xsd:sequence>
      <xsd:element name="pPr"        type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:choice minOccurs="0" maxOccurs="unbounded"/>  <!-- EG_TextRun -->
        <xsd:element name="r"        type="CT_RegularTextRun"/>
        <xsd:element name="br"       type="CT_TextLineBreak"/>
        <xsd:element name="fld"      type="CT_TextField"/>
      </xsd:choice>
      <xsd:element name="endParaRPr" type="CT_TextCharacterProperties" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

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

  <xsd:complexType name="CT_TextCharacterProperties">
    <xsd:sequence>
      <xsd:element name="ln"                   type="CT_LineProperties"         minOccurs="0"/>
      <xsd:group    ref="EG_FillProperties"                                     minOccurs="0"/>
      <xsd:group    ref="EG_EffectProperties"                                   minOccurs="0"/>
      <xsd:element name="highlight"            type="CT_Color"                  minOccurs="0"/>
      <xsd:group    ref="EG_TextUnderlineLine"                                  minOccurs="0"/>
      <xsd:group    ref="EG_TextUnderlineFill"                                  minOccurs="0"/>
      <xsd:element name="latin"                type="CT_TextFont"               minOccurs="0"/>
      <xsd:element name="ea"                   type="CT_TextFont"               minOccurs="0"/>
      <xsd:element name="cs"                   type="CT_TextFont"               minOccurs="0"/>
      <xsd:element name="sym"                  type="CT_TextFont"               minOccurs="0"/>
      <xsd:element name="hlinkClick"           type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="hlinkMouseOver"       type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="rtl"                  type="CT_Boolean"                minOccurs="0"/>
      <xsd:element name="extLst"               type="CT_OfficeArtExtensionList" minOccurs="0"/>
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

  <xsd:complexType name="CT_TextFont">
    <xsd:attribute name="typeface"    type="ST_TextTypeface" use="required"/>
    <xsd:attribute name="panose"      type="s:ST_Panose"/>
    <xsd:attribute name="pitchFamily" type="ST_PitchFamily"  default="0"/>
    <xsd:attribute name="charset"     type="xsd:byte"        default="1"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Hyperlink">
    <xsd:sequence>
      <xsd:element name="snd"    type="CT_EmbeddedWAVAudioFile"   minOccurs="0"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute ref="r:id"/>
      <xsd:attribute name="invalidUrl"     type="xsd:string"  default=""/>
      <xsd:attribute name="action"         type="xsd:string"  default=""/>
      <xsd:attribute name="tgtFrame"       type="xsd:string"  default=""/>
      <xsd:attribute name="tooltip"        type="xsd:string"  default=""/>
      <xsd:attribute name="history"        type="xsd:boolean" default="true"/>
      <xsd:attribute name="highlightClick" type="xsd:boolean" default="false"/>
      <xsd:attribute name="endSnd"         type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_TextCapsType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="small"/>
      <xsd:enumeration value="all"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextFontSize">
    <xsd:restriction base="xsd:int">
      <xsd:minInclusive value="100"/>
      <xsd:maxInclusive value="400000"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextTypeface">
    <xsd:restriction base="xsd:string"/>
  </xsd:simpleType>
