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
