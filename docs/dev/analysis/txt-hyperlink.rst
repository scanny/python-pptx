
Hyperlinks
==========

Overview
--------

PowerPoint supports hyperlinks at two distinct levels:

* a run of text can be a hyperlink
* an entire shape can be a hyperlink

Hyperlinks in PowerPoint can have four types of targets:

* a Internet address, such as https://github/scanny/python-pptx, including an
  optional anchor (e.g. #sub-heading suffix to jump mid-page). This can also
  be an email address, launching the local email client. A mailto: URI is used,
  with subject specifiable using a '?subject=xyz' suffix.

* another file, e.g. another PowerPoint file, including an optional anchor to,
  for example, a specific slide. A file:// URI is used to specify the file
  path.

* another part in the same presentation. This uses an internal relationship
  (in the .rels item) to the target part.

An optional ScreenTip, a roll-over tool-tip sort of message, can also be
specified for any of these link types.

There are some more obscure attributes like "stop playing sound before
navigating" that are available, perhaps meant for kiosk-style applications.

A run or shape can actually have two distinct link behaviors, one for clicking
and another for rolling over with the mouse. These are independent; a run or
shape can have one, the other, both, or neither. These are the two hyperlink
types reported by the Hyperlink.type attribute and using enumeration values
from MsoHyperlinkType.


Minimum viable feature
----------------------

Start with Run since that's the current use case and doesn't require working
out exactly which shapes can be hyperlinked.


.. highlight:: python

::

    r = p.add_run()
    r.text = 'link to python-pptx @ GitHub'
    hlink = r.hyperlink
    hlink.address = 'https://github.com/scanny/python-pptx'


Roadmap items
-------------

* add .hyperlink attribute to Shape and Picture


Resources
---------

* `Hyperlink Object (PowerPoint) on MSDN`_

.. _`Hyperlink Object (PowerPoint) on MSDN`:
   http://msdn.microsoft.com/en-us/library/office/ff746252.aspx


Candidate Protocol
------------------

Add a hyperlink::

    p = shape.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = 'link to python-pptx @ GitHub'
    hlink = r.hyperlink
    hlink.address = 'https://github.com/scanny/python-pptx'

Delete a hyperlink::

    r.hyperlink = None

    # or -----------

    r.hyperlink.address = None  # empty string '' will do it too

A Hyperlink instance is lazy-created on first reference. The object persists
until garbage collected once created. The link XML is not written until
.address is specified. Setting ``hlink.address`` to None or '' causes the
hlink entry to be removed if present.


Candidate API
-------------

_Run.hyperlink

Shape.hyperlink

Hyperlink

  .address - target URL

  .screen_tip - tool-tip text displayed on mouse rollover is slideshow mode

  .type - MsoHyperlinkType (shape or run)

  .show_and_return ...

_Run.rolloverlink would be an analogous property corresponding to the
<a:hlinkMouseOver> element


Enumerations
------------

* MsoHyperlinkType: msoHyperlinkRange, msoHyperlinkShape


Open questions
--------------

* What is the precise scope of shape types that may have a hyperlink applied?
* does leading and trailing space around a hyperlink work as expected?
* not sure what PowerPoint does if you select multiple runs and then insert
  a hyperlink, like including a stretch of bold text surrounded by plain text
  in the selection.


XML specimens
-------------

.. highlight:: xml

Link on overall shape::

    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="2" name="Rectangle 1">
          <a:hlinkClick r:id="rId2"/>
        </p:cNvPr>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      ...
    <p:sp>

Link on a run within a paragraph::

    <a:p>
      <a:r>
        <a:rPr lang="en-US" dirty="0" smtClean="0"/>
        <a:t>Code is available at </a:t>
      </a:r>
      <a:r>
        <a:rPr lang="en-US" dirty="0" smtClean="0">
          <a:hlinkClick r:id="rId2"/>
        </a:rPr>
        <a:t>the python-pptx repository on GitHub</a:t>
      </a:r>
      <a:endParaRPr lang="en-US" dirty="0"/>
    </a:p>


Related Schema Definitions
--------------------------

::

  <xsd:complexType name="CT_TextCharacterProperties">
    <xsd:sequence>
      <xsd:element name="ln"             type="CT_LineProperties"         minOccurs="0"/>
      <xsd:group    ref="EG_FillProperties"                               minOccurs="0"/>
      <xsd:group    ref="EG_EffectProperties"                             minOccurs="0"/>
      <xsd:element name="highlight"      type="CT_Color"                  minOccurs="0"/>
      <xsd:group    ref="EG_TextUnderlineLine"                            minOccurs="0"/>
      <xsd:group    ref="EG_TextUnderlineFill"                            minOccurs="0"/>
      <xsd:element name="latin"          type="CT_TextFont"               minOccurs="0"/>
      <xsd:element name="ea"             type="CT_TextFont"               minOccurs="0"/>
      <xsd:element name="cs"             type="CT_TextFont"               minOccurs="0"/>
      <xsd:element name="sym"            type="CT_TextFont"               minOccurs="0"/>
      <xsd:element name="hlinkClick"     type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="hlinkMouseOver" type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="rtl"            type="CT_Boolean"                minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    ... 19 attributes ...
  </xsd:complexType>

  <xsd:complexType name="CT_Hyperlink">
    <xsd:sequence>
      <xsd:element name="snd"    type="CT_EmbeddedWAVAudioFile"   minOccurs="0"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute ref="r:id"/>     <!-- type="ST_RelationshipId" -->
    <xsd:attribute name="invalidUrl"     type="xsd:string"        default=""/>
    <xsd:attribute name="action"         type="xsd:string"        default=""/>
    <xsd:attribute name="tgtFrame"       type="xsd:string"        default=""/>
    <xsd:attribute name="tooltip"        type="xsd:string"        default=""/>
    <xsd:attribute name="history"        type="xsd:boolean"       default="true"/>
    <xsd:attribute name="highlightClick" type="xsd:boolean"       default="false"/>
    <xsd:attribute name="endSnd"         type="xsd:boolean"       default="false"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_RelationshipId">
    <xsd:restriction base="xsd:string"/>
  </xsd:simpleType>
