.. _BaseSlide:

Base Slide
==========

A slide is the fundamental visual content container in a presentation, that
content taking the form of shape objects. The SlideMaster and SlideLayout
objects are also slides and the three share common behaviors. They each also
have distinctive behaviors. The focus of this page is the common slide
characteristics.


Name
----

For the moment, the only shared attribute is name, stored in the
`p:sld/p:cSld/@name` attribute.

While this attribute is populated by PowerPoint in slide layouts, it is
commonly unpopulated in slides and slide masters. When a slide has no
explicit name, PowerPoint uses the default name 'Slide {n}', where *n* is the
sequence number of the slide in the current presentation. This name is not
written to the `p:cSlide/@name` attribute. In the outline pane, PowerPoint
uses the slide title if there is one.

Slide names may come into things when doing animations. Otherwise they don't
show up in the commonly-used parts of the UI.


XML specimens
-------------

.. highlight:: xml


Example slide contents::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <p:sld
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
    <p:cSld name="Overview">
      <p:spTree>
        ...
      </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
      <a:masterClrMapping/>
    </p:clrMapOvr>
  </p:sld>


Schema excerpt
--------------

.. highlight:: xml

::

  <xsd:element name="sld" type="CT_Slide"/>

  <xsd:complexType name="CT_Slide">  <!-- denormalized -->
    <xsd:sequence minOccurs="1" maxOccurs="1">
      <xsd:element name="cSld"       type="CT_CommonSlideData"/>
      <xsd:element name="clrMapOvr"  type="a:CT_ColorMappingOverride" minOccurs="0"/>
      <xsd:element name="transition" type="CT_SlideTransition"        minOccurs="0"/>
      <xsd:element name="timing"     type="CT_SlideTiming"            minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_ExtensionListModify"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="showMasterSp"     type="xsd:boolean" default="true"/>
    <xsd:attribute name="showMasterPhAnim" type="xsd:boolean" default="true"/>
    <xsd:attribute name="show"             type="xsd:boolean" default="true"/>
  </xsd:complexType>

  <xsd:element name="sldLayout" type="CT_SlideLayout"/>

  <xsd:complexType name="CT_SlideLayout">
    <xsd:sequence>
      <xsd:element name="cSld"       type="CT_CommonSlideData"/>
      <xsd:element name="clrMapOvr"  type="a:CT_ColorMappingOverride" minOccurs="0"/>
      <xsd:element name="transition" type="CT_SlideTransition"        minOccurs="0"/>
      <xsd:element name="timing"     type="CT_SlideTiming"            minOccurs="0"/>
      <xsd:element name="hf"         type="CT_HeaderFooter"           minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_ExtensionListModify"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="showMasterSp"     type="xsd:boolean"        default="true"/>
    <xsd:attribute name="showMasterPhAnim" type="xsd:boolean"        default="true"/>
    <xsd:attribute name="matchingName"     type="xsd:string"         default=""/>
    <xsd:attribute name="type"             type="ST_SlideLayoutType" default="cust"/>
    <xsd:attribute name="preserve"         type="xsd:boolean"        default="false"/>
    <xsd:attribute name="userDrawn"        type="xsd:boolean"        default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_SlideMaster">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="cSld"           type="CT_CommonSlideData"/>
      <xsd:element name="clrMap"         type="a:CT_ColorMapping"/>
      <xsd:element name="sldLayoutIdLst" type="CT_SlideLayoutIdList"     minOccurs="0"/>
      <xsd:element name="transition"     type="CT_SlideTransition"       minOccurs="0"/>
      <xsd:element name="timing"         type="CT_SlideTiming"           minOccurs="0"/>
      <xsd:element name="hf"             type="CT_HeaderFooter"          minOccurs="0"/>
      <xsd:element name="txStyles"       type="CT_SlideMasterTextStyles" minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionListModify"   minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="preserve" type="xsd:boolean" default="false"/>
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

  <xsd:complexType name="CT_ColorMappingOverride">
    <xsd:sequence>
      <xsd:choice minOccurs="1" maxOccurs="1">
        <xsd:element name="masterClrMapping"   type="CT_EmptyElement"/>
        <xsd:element name="overrideClrMapping" type="CT_ColorMapping"/>
      </xsd:choice>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ColorMapping">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bg1"      type="ST_ColorSchemeIndex" use="required"/>
    <xsd:attribute name="tx1"      type="ST_ColorSchemeIndex" use="required"/>
    <xsd:attribute name="bg2"      type="ST_ColorSchemeIndex" use="required"/>
    <xsd:attribute name="tx2"      type="ST_ColorSchemeIndex" use="required"/>
    <xsd:attribute name="accent1"  type="ST_ColorSchemeIndex" use="required"/>
    <xsd:attribute name="accent2"  type="ST_ColorSchemeIndex" use="required"/>
    <xsd:attribute name="accent3"  type="ST_ColorSchemeIndex" use="required"/>
    <xsd:attribute name="accent4"  type="ST_ColorSchemeIndex" use="required"/>
    <xsd:attribute name="accent5"  type="ST_ColorSchemeIndex" use="required"/>
    <xsd:attribute name="accent6"  type="ST_ColorSchemeIndex" use="required"/>
    <xsd:attribute name="hlink"    type="ST_ColorSchemeIndex" use="required"/>
    <xsd:attribute name="folHlink" type="ST_ColorSchemeIndex" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_EmptyElement"/>
