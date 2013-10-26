==================
``CT_Placeholder``
==================

.. highlight:: xml

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 15, 50

   Schema Name  , CT_Placeholder
   Tag(s)       , p:ph
   Namespace    , presentationml (pml.xsd)
   Schema Line  , 1181
   Spec Section , 19.3.1.36


Analysis
========

XPath expression from `<p:sp>` is ``./p:nvSpPr/p:nvPr/p:ph``

I've never seen this element occur in other than a `<p:sp>` element. The
schema indicates it can occur in all the other shape types (group shape,
connector, etc.), but I think that just reflects the inability of XML Schema
to express conditionals, and `<p:ph>` is a child of `<p:nvPr>`, which appears
on shapes of all types.


attributes
^^^^^^^^^^

================  ===  ===================  ==========
name              use  type                 default
================  ===  ===================  ==========
type               ?   ST_PlaceholderType   obj
orient             ?   ST_Direction         horz
sz                 ?   ST_PlaceholderSize   full
idx                ?   xsd:unsignedInt      0
hasCustomPrompt    ?   xsd:boolean          false
================  ===  ===================  ==========


child elements
^^^^^^^^^^^^^^

================  ===  ================================  ========
name               #   type                              line
================  ===  ================================  ========
extLst             ?   CT_ExtensionListModify            775>767
================  ===  ================================  ========


Spec text
^^^^^^^^^

   This element specifies that the corresponding shape should be represented
   by the generating application as a placeholder. When a shape is considered
   a placeholder by the generating application it can have special properties
   to alert the user that they can enter content into the shape. Different
   placeholder types are allowed and can be specified by using the placeholder
   type attribute for this element.


Schema excerpt
^^^^^^^^^^^^^^

::

  <xsd:complexType name="CT_Placeholder">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="type"            type="ST_PlaceholderType" use="optional" default="obj"/>
    <xsd:attribute name="orient"          type="ST_Direction"       use="optional" default="horz"/>
    <xsd:attribute name="sz"              type="ST_PlaceholderSize" use="optional" default="full"/>
    <xsd:attribute name="idx"             type="xsd:unsignedInt"    use="optional" default="0"/>
    <xsd:attribute name="hasCustomPrompt" type="xsd:boolean"        use="optional" default="false"/>
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

