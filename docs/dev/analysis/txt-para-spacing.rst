
Paragraph Spacing
=================

A paragraph has three spacing properties. It can have spacing that separates
it from the prior paragraph, spacing that separates it from the following
paragraph, and line spacing less than or greater than the default single line
spacing.


Candidate protocol
------------------

Get and set space before/after::

    >>> paragraph = TextFrame.add_paragraph()
    >>> paragraph.space_before
    0
    >>> paragraph.space_before = Pt(6)
    >>> paragraph.space_before
    76200
    >>> paragraph.space_before.pt
    6.0

An existing paragraph that somehow has space before or after set in terms of
lines (``<a:spcPct>``) will return a value of Length(0) for that property.

Get and set line spacing::

    >>> paragraph = TextFrame.add_paragraph()
    >>> paragraph.line_spacing
    1.0
    >>> isinstance(paragraph.line_spacing, float)
    True
    >>> paragraph.line_spacing = Pt(18)
    >>> paragraph.line_spacing
    228600
    >>> isinstance(paragraph.line_spacing, Length)
    True
    >>> isinstance(paragraph.line_spacing, int)
    True
    >>> paragraph.line_spacing.pt
    18.0

* The default value for ``Paragraph.line_spacing`` is 1.0 (lines).
* The units of the return value can be distinguished by its type. In
  practice, units can also be distinguished by magnitude. Lines are returned
  as a small-ish float. Point values are returned as a `Length` value.

  If type is the most convenient discriminator, the expression
  `isinstance(line_spacing, float)` will serve as an effective `is_lines`
  predicate.

  In practice, line values will rarely be greater than 3.0 and it would be
  hard to imagine a useful line spacing less than 1 pt (12700). If magnitude
  is the most convenient discriminator, 256 can be assumed to be a safe
  threshold value.


MS API
------

LineRuleAfter
    Determines whether line spacing after the last line in each paragraph is
    set to a specific number of points or lines. Read/write.

LineRuleBefore
    Determines whether line spacing before the first line in each paragraph
    is set to a specific number of points or lines. Read/write.

LineRuleWithin
    Determines whether line spacing between base lines is set to a specific
    number of points or lines. Read/write.

SpaceAfter
    Returns or sets the amount of space after the last line in each paragraph
    of the specified text, in points or lines. Read/write.

SpaceBefore
    Returns or sets the amount of space before the first line in each
    paragraph of the specified text, in points or lines. Read/write.

SpaceWithin
    Returns or sets the amount of space between base lines in the specified
    text, in points or lines. Read/write.


PowerPoint behavior
-------------------

* The UI doesn't appear to allow line spacing to be set in units of lines,
  although the XML and MS API allow it to be specified so.


XML specimens
-------------

.. highlight:: xml

Default paragraph spacing::

  <a:p>
    <a:r>
      <a:t>Paragraph with default spacing</a:t>
    </a:r>
  </a:p>

Spacing before = 6 pt::

  <a:p>
    <a:pPr>
      <a:spcBef>
        <a:spcPts val="600"/>
      </a:spcBef>
    </a:pPr>
    <a:r>
      <a:t>Paragraph spacing before = 6pt</a:t>
    </a:r>
  </a:p>

Spacing after = 12 pt::

  <a:p>
    <a:pPr>
      <a:spcAft>
        <a:spcPts val="1200"/>
      </a:spcAft>
    </a:pPr>
    <a:r>
      <a:t>Paragraph spacing after = 12pt</a:t>
    </a:r>
  </a:p>

Line spacing = 24 pt::

  <a:p>
    <a:pPr>
      <a:lnSpc>
        <a:spcPts val="2400"/>
      </a:lnSpc>
    </a:pPr>
    <a:r>
      <a:t>Paragraph line spacing = 24pt</a:t>
    </a:r>
    <a:br/>
    <a:r>
      <a:t>second line</a:t>
    </a:r>
  </a:p>

Line spacing = 2 lines::

  <a:p>
    <a:pPr>
      <a:lnSpc>
        <a:spcPct val="200000"/>
      </a:lnSpc>
    </a:pPr>
    <a:r>
      <a:t>Paragraph line spacing = 2 line</a:t>
    </a:r>
    <a:br/>
    <a:r>
      <a:t>second line</a:t>
    </a:r>
  </a:p>


Schema excerpt
--------------

.. highlight:: xml

::

  <xsd:complexType name="CT_TextParagraph">
    <xsd:sequence>
      <xsd:element name="pPr"        type="CT_TextParagraphProperties" minOccurs="0"/>
      <xsd:group    ref="EG_TextRun" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="endParaRPr" type="CT_TextCharacterProperties" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextParagraphProperties">
    <xsd:sequence>
      <xsd:element name="lnSpc"       type="CT_TextSpacing"             minOccurs="0"/>
      <xsd:element name="spcBef"      type="CT_TextSpacing"             minOccurs="0"/>
      <xsd:element name="spcAft"      type="CT_TextSpacing"             minOccurs="0"/>
      <xsd:choice minOccurs="0">       <!-- EG_TextBulletColor -->
        <xsd:element name="buClrTx"   type="CT_TextBulletColorFollowText"/>
        <xsd:element name="buClr"     type="CT_Color"/>
      </xsd:choice>
      <xsd:choice minOccurs="0">       <!-- EG_TextBulletSize -->
        <xsd:element name="buSzTx"    type="CT_TextBulletSizeFollowText"/>
        <xsd:element name="buSzPct"   type="CT_TextBulletSizePercent"/>
        <xsd:element name="buSzPts"   type="CT_TextBulletSizePoint"/>
      </xsd:choice>
      <xsd:choice minOccurs="0">       <!-- EG_TextBulletTypeface -->
        <xsd:element name="buFontTx"  type="CT_TextBulletTypefaceFollowText"/>
        <xsd:element name="buFont"    type="CT_TextFont"/>
      </xsd:choice>
      <xsd:choice minOccurs="0">       <!-- EG_TextBullet -->
        <xsd:element name="buNone"    type="CT_TextNoBullet"/>
        <xsd:element name="buAutoNum" type="CT_TextAutonumberBullet"/>
        <xsd:element name="buChar"    type="CT_TextCharBullet"/>
        <xsd:element name="buBlip"    type="CT_TextBlipBullet"/>
      </xsd:choice>
      <xsd:element name="tabLst"      type="CT_TextTabStopList"         minOccurs="0"/>
      <xsd:element name="defRPr"      type="CT_TextCharacterProperties" minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_OfficeArtExtensionList"  minOccurs="0"/>
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

  <xsd:complexType name="CT_TextSpacing">
    <xsd:choice>
      <xsd:element name="spcPct" type="CT_TextSpacingPercent"/>
      <xsd:element name="spcPts" type="CT_TextSpacingPoint"/>
    </xsd:choice>
  </xsd:complexType>

  <xsd:complexType name="CT_TextSpacingPercent">
    <xsd:attribute name="val" type="ST_TextSpacingPercentOrPercentString" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_TextSpacingPercentOrPercentString">
    <xsd:union memberTypes="ST_TextSpacingPercent s:ST_Percentage"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TextSpacingPercent">
    <xsd:restriction base="ST_PercentageDecimal">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="13200000"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Percentage">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="-?[0-9]+(\.[0-9]+)?%"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_TextSpacingPoint">
    <xsd:attribute name="val" type="ST_TextSpacingPoint" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_TextSpacingPoint">
    <xsd:restriction base="xsd:int">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="158400"/>
    </xsd:restriction>
  </xsd:simpleType>
