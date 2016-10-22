.. _placeholder:

Placeholders
============

.. toctree::
   :titlesonly:

   slide-placeholders/index
   layout-placeholders
   master-placeholders

From a user perspective, a placeholder is a container into which content can
be inserted. The position, size, and formatting of content inserted into the
placeholder is determined by the placeholder, allowing those key presentation
design characteristics to be determined at design time and kept consistent
across presentations created from a particular template.

The placeholder appearing on a slide is only part of the overall placeholder
mechanism however. Placeholder behavior requires three different categories
of placeholder shape; those that exist on a slide master, those on a slide
layout, and those that ultimately appear on a slide in a presentation.

These three categories of placeholder participate in a property inheritance
hierarchy, either as an inheritor, an inheritee, or both. Placeholder shapes
on masters are inheritees only. Conversely placeholder shapes on slides are
inheritors only. Placeholders on slide layouts are both, a possible inheritor
from a slide master placeholder and an inheritee to placeholders on slides
linked to that layout.

A layout inherits from its master differently than a slide inherits from
its layout. A layout placeholder inherits from the master placeholder sharing
the same type. A slide placeholder inherits from the layout placeholder
having the same `idx` value.

Glossary
--------

placeholder shape
    A shape on a slide that inherits from a layout placeholder.

layout placeholder
    a shorthand name for the placeholder shape on the slide layout from which
    a particular placeholder on a slide inherits shape properties

master placeholder
    the placeholder shape on the slide master which a layout placeholder
    inherits from, if any.


Schema excerpt
--------------

.. highlight:: xml

::

  <xsd:complexType name="CT_ApplicationNonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element   name="ph"            type="CT_Placeholder"      minOccurs="0"/>
      <xsd:choice minOccurs="0">  <!-- a:EG_Media -->
        <xsd:element name="audioCd"       type="CT_AudioCD"/>
        <xsd:element name="wavAudioFile"  type="CT_EmbeddedWAVAudioFile"/>
        <xsd:element name="audioFile"     type="CT_AudioFile"/>
        <xsd:element name="videoFile"     type="CT_VideoFile"/>
        <xsd:element name="quickTimeFile" type="CT_QuickTimeFile"/>
      </xsd:choice>
      <xsd:element   name="custDataLst"   type="CT_CustomerDataList" minOccurs="0"/>
      <xsd:element   name="extLst"        type="CT_ExtensionList"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="isPhoto"   type="xsd:boolean" default="false"/>
    <xsd:attribute name="userDrawn" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Placeholder">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="type"            type="ST_PlaceholderType" default="obj"/>
    <xsd:attribute name="orient"          type="ST_Direction"       default="horz"/>
    <xsd:attribute name="sz"              type="ST_PlaceholderSize" default="full"/>
    <xsd:attribute name="idx"             type="xsd:unsignedInt"    default="0"/>
    <xsd:attribute name="hasCustomPrompt" type="xsd:boolean"        default="false"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_PlaceholderType">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="title"/>
      <xsd:enumeration value="body"/>
      <xsd:enumeration value="ctrTitle"/>
      <xsd:enumeration value="subTitle"/>
      <xsd:enumeration value="dt"/>
      <xsd:enumeration value="sldNum"/>
      <xsd:enumeration value="ftr"/>
      <xsd:enumeration value="hdr"/>
      <xsd:enumeration value="obj"/>
      <xsd:enumeration value="chart"/>
      <xsd:enumeration value="tbl"/>
      <xsd:enumeration value="clipArt"/>
      <xsd:enumeration value="dgm"/>
      <xsd:enumeration value="media"/>
      <xsd:enumeration value="sldImg"/>
      <xsd:enumeration value="pic"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_PlaceholderSize">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="full"/>
      <xsd:enumeration value="half"/>
      <xsd:enumeration value="quarter"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Direction">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="horz"/>
      <xsd:enumeration value="vert"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_PlaceholderSize">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="full"/>
      <xsd:enumeration value="half"/>
      <xsd:enumeration value="quarter"/>
    </xsd:restriction>
  </xsd:simpleType>


Notes on footer item placeholders
---------------------------------

Recommendation and Argumentation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Place page number (and perhaps copyright, date) on master in plain text boxes.
Use "Hide Background Graphics" option on title slide to prevent them showing
there if desired.

* In the slide placeholder method, an individually selectable footer item
  placeholder is present on the slide:

  + It can be moved accidentally, losing the placement consistency originally
    imposed by the master. Recovering exact placement requires reapplying the
    layout which may make other, unwanted changes.
  + It can be deleted, which could be a plus or minus depending.
  + It can be formatted inconsistently.
  + In can be accidentally selected and generally pollutes the shape space of
    the slide.

* I'm not sure python-pptx does the right thing when adding slides from layouts
  containing footer items, at least not the same thing as PowerPoint does in
  automatically working out which footer items to copy over.

  Also, there's no method to copy over footer items from the layout.

* Alternatives:

  1. A page number can be placed on the master as a text box. This causes it to
     appear on every layout and every slide. No refreshing is necessary as this
     item is inherited, not copied at slide or layout create time.

     It can be excluded from an individual layout (such as a title slide
     layout) by setting the `Hide Background Graphics` flag on
     a layout-by-layout basis. Unfortunately, this also hides other items such
     as a logo, copyright text, and slide background, but these can be
     reapplied individually directly to that slide layout. Since the title
     slide layout is generally the only one folks want to hide a slide number
     from, this is perhaps not a great burden. Others, like perhaps closing
     slides, might have their own background anyway.


Hypothesis:
~~~~~~~~~~~

* There is a class of placeholders known as "footer elements", comprising date
  and time, footer text, and slide number.

* These placeholders can only be directly added to a master. They are not
  available as options to add directly to a layout or to a slide.

* In every case, these placeholders are copied, although some attributes may be
  inherited at run time (not sure). Footer items are copied from the master to
  a new layout. Likewise, items are copied from a layout to a slide, but only
  when explicitly asked for using Insert > Header and Footer. PowerPoint
  sometimes adds these automatically when they're already present on other
  slides (don't know the rules it uses).


Intended procedure
~~~~~~~~~~~~~~~~~~

This is how slide footer items are designed to be used:

* Add footer placeholders to master with intended standard position and
  formatting.

* Available by default on each layout. Can be disabled for individual layouts,
  for example, the title slide layout using the Slide Master Ribbon > Edit
  Layout > Allow Footers checkbox. Also can be deleted and formatting changed
  individually on a layout-by-layout basis.

* Hidden by default on new slides. Use Insert > Header and Footer dialog to add
  which items to show. Can be applied to all slides. Under certain
  circumstances, PowerPoint adds selected footer items to new slides
  automatically, based on them being present on other slides already in the
  presentation.


PowerPoint Behaviors
~~~~~~~~~~~~~~~~~~~~

* Layout: Slide Master (ribbon) > Edit Layout > Allow Footers (checkbox)

  + Adding or removing a footer item from a layout, does not automatically
    reflect that change on existing slides based on that layout. The change is
    also not reflected when you reapply the layout, at least for deleting case.

    - It does not appear when you reapply the layout ('Reset Layout to Default
      Settings' on Mac)

* Position appears to be inherited as long as you haven't moved the placeholder
  on the slide. This applies at master -> layout as well as layout -> slide
  level.

  + Refreshing the layout after the page number item has been deleted leaves
    the page number item, but repositions it to a default (right justified to
    right slide edge) position. Apparently it attempts to inherit the position
    from the layout but not presence (or not).

* The `Allow Footers` checkbox is inactive (not greyed out but can't be set on)
  when there are no footer placeholder items on the master.

* Allow footers seems to be something of a one-way street. You can't clear the
  checkbox once it's set. However, if you delete the footer placeholder(s), the
  checkbox clears all by itself.

* The placeholder copied from master to layout can be repositioned and
  formatted (font etc.).

* Menu > Insert > Header and Footer ... Is available on individual slides and
  can be applied to all slides. It also works for removing footer items. This
  applies to *slides*, not to slide layouts.

  This dialog is also available from the Page Setup dialog.

* Hide background graphics flag does not control footer items. It does however
  control visibility of all non-placeholder items on the layout (or slide).
  It's available at both the layout and slide level. Not sure how that's
  reflected in the XML.

* When creating a new slide layout, the footer placeholders from the master are
  all automatically shown. This is to say that the "Allow Footers" ribbon
  option defaults to True when adding a new slide layout.

  Clicking Allow Footers off hides all the footer items. They can also be
  deleted or formatted individually.

* In order to appear (at create time anyway), footer items need to be:

  1. Present on the master
  2. Present on the layout ('Allow Footers' turned on, ph not deleted)
  3. Insert Header and Footer turned on for each ph element to be shown. This
     is applied on a slide-by-slide basis.


Notes
~~~~~

* Header is a misnomer as far as slides are concerned. There are no header
  items for a slide. A notes page can have a header though, which I expect is
  why it appears, since the same dialog is used for slides and notes pages (and
  handouts).

* A master is intended only to contain certain items:

  + Title
  + Body
  + Footer placeholders (date/time, footer text, slide number)
  + Background and theme items, perhaps logos and static text


* http://www.powerpointninja.com/templates/powerpoint-templates-beware-of-the-footers/

* https://github.com/scanny/python-pptx/issues/64
