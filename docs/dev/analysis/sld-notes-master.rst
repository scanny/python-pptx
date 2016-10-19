
Notes Master
============

A presentation may have a notes master part (zero or one). The notes master
determines the background elements and default layout of a notes page.

In the PowerPoint UI, a notes item is called a "notes page"; internally
however, it is referred to as a "notes slide", is based on the Slide object,
and shares many behaviors with it such as containing shapes and inheriting
from a master.

A notes slide is created on-first use, perhaps most frequently by typing text
into the notes pane. When a new notes slide is created, it is based on the
notes master. If no notes master is yet present, a new one is created from
an internal PowerPoint-preset default and added to the package.

When a notes slide is created from the notes master, certain placeholders are
partially "cloned" from the master. All other placeholders and shapes remain
only on the master. The cloning process only creates placeholder shapes on
the notes slide, but does not copy position or size information or other
formatting. By default, position, size, and other formatting is inherited
from the corresponding master placeholder. This achieves consistency through
a presentation.

If position, size, etc. is changed on a NotesSlide, the new position
overrides that on the master. This override is property-by-property.

A notes slide maintains a relationship with the master it was created from.
This relationship is traversed to access the master placeholder properties
for inheritance purposes.

A master can contain at most one each of six different placeholders: slide
image, notes (textbox), header, footer, date, and slide number. The header,
footer, and date placeholders are not copied to the notes slide by default.
These three placeholders can be individually shown or hidden (copied to or
deleted from the notes slide actually) using the Insert > Header and
Footer... menu option/dialog. This dialog is also available from the Page
Setup dialog.

All other items on the notes master, such as background color/texture/image
and any additional text boxes or other shapes such as a logo, constitute
"background graphics" and are shown by default. The may be hidden or re-shown
on a notes slide-by-slide basis using the Format Background... (context or
menu bar) menu item.


Candidate Protocol
------------------

A NotesMaster part is created on first access when not yet present::

    >>> notes_master = presentation.notes_master
    >>> notes_master
    <pptx.slide.NotesMaster object at 0x10698c1e0>

It provides access to its placeholders and its shapes (which include
placeholders)::

    >>> notes_master.shapes
    <pptx.shapes.shapetree.MasterShapes object at 0x10698d140>
    >>> notes_master.placeholders
    <pptx.shapes.shapetree.MasterPlaceholders object at 0x1069902f0>

These placeholders and other shapes can be manipulated as usual.


Understanding Placeholders implementation
-----------------------------------------

* SlidePlaceholders is a direct implementation, not subclassing Shapes

Call stacks:
------------

* _BaseSlide.placeholders
    => BasePlaceholders()
        => _BaseShapes + _is_member_elm() override

* NotesMaster.placeholders
    => _BaseMaster.placeholders
        => MasterPlaceholders()
            => BasePlaceholders() + get() + _shape_factory() override
                => _BaseShapes + _is_member_elm() override


PowerPoint behavior
-------------------

* A NotesSlide placeholder inherits properties from the NotesMaster
  placeholder *of the same type* (there can be at most one).

* For some reason PowerPoint adds a second theme for the NotesMaster, not
  sure what that's all about. I'll just add the default theme2.xml, then
  folks can avoid that if there are problems just by pre-loading
  a notesMaster of their own in their template.


Related Schema Definitions
--------------------------

.. highlight:: xml

The root element of a notesSlide part is a `p:notes` element::

  <xsd:element name="notesMaster" type="CT_NotesMaster"/>

  <xsd:complexType name="CT_NotesMaster">
    <xsd:sequence>
      <xsd:element name="cSld"       type="CT_CommonSlideData"/>
      <xsd:element name="clrMap"     type="a:CT_ColorMapping"/>
      <xsd:element name="hf"         type="CT_HeaderFooter"        minOccurs="0"/>
      <xsd:element name="notesStyle" type="a:CT_TextListStyle"     minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_CommonSlideData">
    <xsd:sequence>
      <xsd:element name="bg"          type="CT_Background"       minOccurs="0"/>
      <xsd:element name="spTree"      type="CT_GroupShape"/>
      <xsd:element name="custDataLst" type="CT_CustomerDataList" minOccurs="0"/>
      <xsd:element name="controls"    type="CT_ControlList"      minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="name" type="xsd:string" use="optional" default=""/>
  </xsd:complexType>

  <xsd:element name="theme" type="CT_OfficeStyleSheet"/>

  <xsd:complexType name="CT_OfficeStyleSheet">
    <xsd:sequence>
      <xsd:element name="themeElements"     type="CT_BaseStyles"/>
      <xsd:element name="objectDefaults"    type="CT_ObjectStyleDefaults"    minOccurs="0"/>
      <xsd:element name="extraClrSchemeLst" type="CT_ColorSchemeList"        minOccurs="0"/>
      <xsd:element name="custClrLst"        type="CT_CustomColorList"        minOccurs="0"/>
      <xsd:element name="extLst"            type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="name" type="xsd:string" default=""/>
  </xsd:complexType>
