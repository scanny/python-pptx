
Slide Layout
============

A slide layout acts as an property inheritance base for zero or more slides.
This provides a certain amount of separation between formatting and content
and contributes to visual consistency across the slides of a presentation.


Schema excerpt
--------------

.. highlight:: xml

::

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
