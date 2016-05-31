
Autofit setting
===============

Overview
--------

PowerPoint provides the *autofit* property to specify how to handle text that
is too big to fit within its shape. The three possible settings are:

* resize shape to fit text
* resize text to fit shape
* resize neither the shape nor the text, allowing the text to extend beyond
  the bounding box

Auto size is closely related to word wrap.

There are certain constraining settings for resizing the text, not sure if
those are supported in PowerPoint or not.


Protocol
--------

::

    >>> text_frame = shape.text_frame
    >>> text_frame.auto_size
    None
    >>> text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    >>> text_frame.auto_size
    1
    >>> text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    >>> text_frame.auto_size
    2
    >>> str(text_frame.auto_size)
    'MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE'
    >>> text_frame.auto_size = MSO_AUTO_SIZE.NONE
    >>> text_frame.auto_size
    0
    >>> text_frame.auto_size = None
    >>> text_frame.auto_size
    None


Scenarios
---------

**Default new textbox.**  When adding a new texbox from the PowerPoint
toolbar, the new textbox sizes itself to fit the entered text. This
corresponds to the setting MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT and ``word_wrap
= None``.

``MSO_AUTO_SIZE.NONE``, ``word_wrap = False``. Text is free form, with
PowerPoint exhibiting no formatting behavior. Text appears just as entered
and lines are broken only where hard breaks are inserted.

``MSO_AUTO_SIZE.NONE``, ``word_wrap = True``. Text is wrapped into
a column the width of the shape, but is not constrained vertically.

``MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT``, ``word_wrap = False``. The width of
the shape expands as new text is entered. Line breaks occur only where hard
breaks are entered.  The height of the shape grows to accommodate the number
of lines of entered text.  Width and height shrink as extents of the text are
reduced by deleting text or reducing font size.

``MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT``, ``word_wrap = True``. Text is
wrapped into a column the width of the shape, with the height of the shape
growing to accommodate the resulting number of lines.

``MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE``, ``word_wrap = False``. Experiment
...

``MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE``, ``word_wrap = True``. Experiment ...


XML specimens
-------------

.. highlight:: xml

Default new textbox::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="4" name="Foobar 3"/>
      <p:cNvSpPr txBox="1"/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="914400" y="914400"/>
        <a:ext cx="914400" cy="914400"/>
      </a:xfrm>
      <a:prstGeom prst="rect">
        <a:avLst/>
      </a:prstGeom>
      <a:noFill/>
    </p:spPr>
    <p:txBody>
      <a:bodyPr wrap="none">
        <a:spAutoFit/>
      </a:bodyPr>
      <a:lstStyle/>
      <a:p/>
    </p:txBody>
  </p:sp>


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

  <xsd:complexType name="CT_TextBodyProperties">
    <xsd:sequence>
      <xsd:element name="prstTxWarp" type="CT_PresetTextShape"        minOccurs="0"/>
      <xsd:group   ref="EG_TextAutofit"                               minOccurs="0"/>
      <xsd:element name="scene3d"    type="CT_Scene3D"                minOccurs="0"/>
      <xsd:group   ref="EG_Text3D"                                    minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>

  <xsd:group name="EG_TextAutofit">
    <xsd:choice>
      <xsd:element name="noAutofit"   type="CT_TextNoAutofit"/>
      <xsd:element name="normAutofit" type="CT_TextNormalAutofit"/>
      <xsd:element name="spAutoFit"   type="CT_TextShapeAutofit"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_TextNormalAutofit">
    <xsd:attribute name="fontScale" type="ST_TextFontScalePercentOrPercentString"
                   use="optional" default="100%"/>
    <xsd:attribute name="lnSpcReduction" type="ST_TextSpacingPercentOrPercentString"
                   use="optional" default="0%"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TextShapeAutofit"/>

  <xsd:complexType name="CT_TextNoAutofit"/>

  <xsd:group name="EG_Text3D">
    <xsd:choice>
      <xsd:element name="sp3d"   type="CT_Shape3D"/>
      <xsd:element name="flatTx" type="CT_FlatText"/>
    </xsd:choice>
  </xsd:group>
