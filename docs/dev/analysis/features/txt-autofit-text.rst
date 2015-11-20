
Text - Auto-fit text to shape
=============================

An AutoShape has a text frame, referred to in the PowerPoint UI as the
shape's *Text Box*. One of the settings provided is *Autofit*, which can be
one of "Do not autofit", "Resize text to fit shape", or "Resize shape to fit
text". The scope of this analysis is how best to implement the "Resize text
to fit shape" behavior.

A robust implementation would be complex and would lead the project outside
the currently intended scope. In particular, because the shape size, text
content, the "full" point size of the text, and the autofit and wrap settings
of the textframe all interact to determine the proper "fit" of adjusted text,
all events that could change the state of any of these five factors would
need to be coupled to an "update" method. There would also need to be at
least two "fitting" algorithms, one for when wrap was turned on and another
for when it was off.

The initial solution we'll pursue is a special-purpose method on TextFrame
that reduces the permutations of these variables to one and places
responsibility on the developer to call it at the appropriate time. The
method will calculate based on the current text and shape size, set wrap on,
and set auto_size to fit text to the shape. If any of the variables change,
the developer will be responsible for re-calling the method.


Current constraints
-------------------

:meth:`.TextFrame.autofit_text`
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

* User must manually set all shape text to a uniform 12pt default full-size
  font point size.

  + This is intended to be done before the call, but might work okay if done
    after too, as long as it matches the default 12pt.

* Only 12pt "full-size" is supported. There is no mechanism to specify other
  sizes.
* Unconditionally sets autofit and wrap.
* Path to font file must be provided manually.
* User must manually set the font typeface of all shape text to match the
  provided font file.


Incremental enhancements
------------------------

* **.fit_text() or .autofit_text().** Two related methods are used to fit
  text in a shape using different approaches. ``TextFrame.autofit_text()``
  uses the `TEXT_TO_FIT_SHAPE` autofit setting to shrink a full-size
  paragraph of text. Later edits to that text using PowerPoint will re-fit
  the text automatically, up to the original (full-size) font size.

  ``TextFrame.fit_text()`` takes the approach of simply setting the font size
  for all text in the shape to the best-fit size. No automatic resizing
  occurs on later edits in PowerPoint, although the user can switch on
  auto-fit for that text box, perhaps after setting the full-size point size
  to their preferred size.

* **Specified full point size.** Allow the point size maximum to be
  specified; defaults to 18pt. All text in the shape is set to this size
  before calculating the best-fit based on that size.

* **Specified font.** In the process, this specifies the font to use,
  although it may require a tuple specifying the type family name as well as
  the bold and italic states, plus a file path.

* **Auto-locate installed font file.** Search for and use the appropriate
  locally-installed font file corresponding to the selected typeface. On
  Windows, the font directory can be located using a registry key, and is
  perhaps often `C:\Windows\Fonts`. However the font filename is not the same
  as the UI typeface name, so some mapping would be required, including
  detecting whether bold and/or italic were specified.

* **Accommodate line breaks.** Either from `<a:br>` elements or multiple
  paragraphs. Would involve ending lines at a break, other than the last
  paragraph.

* **Add line space reduction.** PowerPoint always reduces line spacing by as
  much as 20% to maximize the font size used. Add this factor into the
  calculation to improve the exact match likelihood for font scale and
  lnSpcReduction and thereby reduce occurence of "jumping" of text to a new
  size on edit.


Candidate Protocol
------------------

Shape size and text are set before calling
:meth:`.TextFrame.autofit_text` or :meth:`.TextFrame.fit_text`::

    >>> shape.width, shape.height = cx, cy
    >>> shape.text = 'Lorem ipsum .. post facto.'
    >>> text_frame = shape.text_frame
    >>> text_frame.auto_size, text_frame.word_wrap, text_frame._font_scale
    (None, False, None)

Calling :meth:`TextFrame.autofit_text` turns on auto fit (text to shape),
switches on word wrap, and calculates the best-fit font scaling factor::

    >>> text_frame.autofit_text()
    >>> text_frame.auto_size, text_frame.word_wrap, text_frame._font_scale
    (TEXT_TO_FIT_SHAPE (2), True, 55.0)

Calling :meth:`TextFrame.fit_text` produces the same sized text, but autofit
is not turned on. Rather the actual font size for text in the shape is set to
the calculated "best fit" point size::

    >>> text_frame.fit_text()
    >>> text_frame.auto_size, text_frame.word_wrap, text_frame._font_scale
    (None, True, None)
    >>> text_frame.paragraphs[0].font.size.pt
    10


Proposed |pp| behavior
----------------------

The :meth:`TextFrame.fit_text` method produces the following side effects:

* :attr:`TextFrame.auto_size` is set to
  :attr:`MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`
* :attr:`TextFrame.word_wrap` is set to |True|.
* A suitable value is calculated for `<a:normAutofit fontScale="?"/>`. The
  `fontScale` attribute is set to this value and the `lnSpcReduction`
  attribute is removed, if present.

The operation can be undone by assigning |None|, :attr:`MSO_AUTO_SIZE.NONE`,
or :attr:`MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT` to `TextFrame.auto_size`.


PowerPoint behavior
-------------------

* PowerPoint shrinks text in whole-number font sizes.

* The behavior interacts with *Wrap text in shape*. The behavior we want here
  is only when wrap is turned on. When wrap is off, only height and manual
  line breaks are taken into account. Long lines simply extend outside the
  box.

* When assigning a font size to a shape, PowerPoint applies that font size at
  the run level, adding a `sz` attribute to the `<a:rPr>` element for every
  content child of every `<a:p>` element in the shape. The sentinel
  `<a:endParaRPr>` element also gets a `sz` attribute set to that size, but
  only in the last paragraph, it appears.


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
auto-reduced to 10pt. ``<a:t>`` element text wrapped for compact display::

  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0">
      <a:normAutofit fontScale="55000" lnSpcReduction="20000"/>
    </a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" dirty="0" smtClean="0"/>
        <a:t>The art and craft of designing typefaces is called type design.
             Designers of typefaces are called type designers and are often
             employed by type foundries. In digital typography, type
             designers are sometimes also called font developers or font
             designers.</a:t>
      </a:r>
      <a:endParaRPr lang="en-US" dirty="0"/>
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

  <xsd:complexType name="CT_TextNormalAutofit">
    <xsd:attribute name="fontScale" type="ST_TextFontScalePercentOrPercentString"
                   use="optional" default="100%"/>
    <xsd:attribute name="lnSpcReduction" type="ST_TextSpacingPercentOrPercentString"
                   use="optional" default="0%"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TextShapeAutofit"/>

  <xsd:complexType name="CT_TextNoAutofit"/>

  <xsd:simpleType name="ST_TextFontScalePercentOrPercentString">
    <xsd:union memberTypes="ST_TextFontScalePercent s:ST_Percentage"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextFontScalePercent">
    <xsd:restriction base="ST_PercentageDecimal">
      <xsd:minInclusive value="1000"/>
      <xsd:maxInclusive value="100000"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Percentage">  <!-- s:ST_Percentage -->
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="-?[0-9]+(\.[0-9]+)?%"/>
    </xsd:restriction>
  </xsd:simpleType>
