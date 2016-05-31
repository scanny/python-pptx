########################
Core Document Properties
########################

:Updated:  2013-06-22
:Author:   Steve Canny
:Status:   **WORKING DRAFT**


Introduction
============

The 'Core' in core document properties refers to `Dublin Core`_, a metadata
standard that defines a core set of elements to describe resources.

ISO/IEC 29500-2 Section 11


API Sketch
==========

Presentation.core_properties

_CoreProperties():

All string values are limited to 255 chars (unicode chars, not bytes)

* title
* subject
* author
* keywords
* comments
* last_modified_by
* revision (integer)
* created (date in format: '2013-04-06T06:03:36Z')
* modified (date in same format)
* category


Status and company (and many others) are custom properties, held in
``app.xml``.


XML produced by PowerPoint® client
==================================

.. highlight:: xml

::

    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
      <dc:title>Core Document Properties Exploration</dc:title>
      <dc:subject>PowerPoint core document properties</dc:subject>
      <dc:creator>Steve Canny</dc:creator>
      <cp:keywords>powerpoint; open xml; dublin core; microsoft office</cp:keywords>
      <dc:description>One thing I'd like to discover is just how line wrapping is handled in the comments. This paragraph is all on a single line._x000d__x000d_This is a second paragraph separated from the first by two line feeds.</dc:description>
      <cp:lastModifiedBy>Steve Canny</cp:lastModifiedBy>
      <cp:revision>2</cp:revision>
      <dcterms:created xsi:type="dcterms:W3CDTF">2013-04-06T06:03:36Z</dcterms:created>
      <dcterms:modified xsi:type="dcterms:W3CDTF">2013-06-15T06:09:18Z</dcterms:modified>
      <cp:category>analysis</cp:category>
    </cp:coreProperties>


Schema
======

::

    <xs:schema
      targetNamespace="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
      xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
      xmlns:xs="http://www.w3.org/2001/XMLSchema"
      xmlns:dc="http://purl.org/dc/elements/1.1/"
      xmlns:dcterms="http://purl.org/dc/terms/"
      elementFormDefault="qualified"
      blockDefault="#all">

      <xs:import
        namespace="http://purl.org/dc/elements/1.1/"
        schemaLocation="http://dublincore.org/schemas/xmls/qdc/2003/04/02/dc.xsd"/>
      <xs:import
        namespace="http://purl.org/dc/terms/"
        schemaLocation="http://dublincore.org/schemas/xmls/qdc/2003/04/02/dcterms.xsd"/>
      <xs:import
        id="xml"
        namespace="http://www.w3.org/XML/1998/namespace"/>

      <xs:element name="coreProperties" type="CT_CoreProperties"/>

      <xs:complexType name="CT_CoreProperties">
        <xs:all>
          <xs:element name="category"        type="xs:string"   minOccurs="0"/>
          <xs:element name="contentStatus"   type="xs:string"   minOccurs="0"/>
          <xs:element ref="dcterms:created"                     minOccurs="0"/>
          <xs:element ref="dc:creator"                          minOccurs="0"/>
          <xs:element ref="dc:description"                      minOccurs="0"/>
          <xs:element ref="dc:identifier"                       minOccurs="0"/>
          <xs:element name="keywords"        type="CT_Keywords" minOccurs="0"/>
          <xs:element ref="dc:language"                         minOccurs="0"/>
          <xs:element name="lastModifiedBy"  type="xs:string"   minOccurs="0"/>
          <xs:element name="lastPrinted"     type="xs:dateTime" minOccurs="0"/>
          <xs:element ref="dcterms:modified"                    minOccurs="0"/>
          <xs:element name="revision"        type="xs:string"   minOccurs="0"/>
          <xs:element ref="dc:subject"                          minOccurs="0"/>
          <xs:element ref="dc:title"                            minOccurs="0"/>
          <xs:element name="version"         type="xs:string"   minOccurs="0"/>
        </xs:all>
      </xs:complexType>

      <xs:complexType name="CT_Keywords" mixed="true">
        <xs:sequence>
          <xs:element name="value" minOccurs="0" maxOccurs="unbounded" type="CT_Keyword"/>
        </xs:sequence>
        <xs:attribute ref="xml:lang" use="optional"/>
      </xs:complexType>

      <xs:complexType name="CT_Keyword">
        <xs:simpleContent>
          <xs:extension base="xs:string">
            <xs:attribute ref="xml:lang" use="optional"/>
          </xs:extension>
        </xs:simpleContent>
      </xs:complexType>

    </xs:schema>



.. _Dublin Core:
   http://en.wikipedia.org/wiki/Dublin_Core
