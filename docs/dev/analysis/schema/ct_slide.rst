===================
``CT_Slide``
===================

.. highlight:: xml

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 15, 50

   Schema Name  , CT_Slide
   Spec Name    , Presentation Slide
   Tag(s)       , p:sld
   Namespace    , presentationml (pml.xsd)
   Schema Line  , 1344
   Spec Section , 19.3.1.38


Analysis
========

The ``<p:sld>`` element is the root element for a slide part.


Spec text
=========

   This element specifies a slide within a slide list. The slide list is used
   to specify an ordering of slides.
   
   (*sc: this text appears to be in error, being an element in a sldLst is*
   *perhaps one role, but the main role I expect is its role as the root*
   *element of a slide part*.)


Minimal Slide part
==================

The following XML element represents a minimal slide part::

   <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
   <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
     <p:cSld>
       <p:spTree>
         <p:nvGrpSpPr>
           <p:cNvPr id="1" name=""/>
           <p:cNvGrpSpPr/>
           <p:nvPr/>
         </p:nvGrpSpPr>
         <p:grpSpPr/>
       </p:spTree>
     </p:cSld>
   </p:sld>


Schema excerpt
==============

::

  <xsd:element name="sld" type="CT_Slide"/>

  <xsd:complexType name="CT_Slide">  <!-- denormalized -->
    <xsd:sequence>
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

  <xsd:complexType name="CT_CommonSlideData">
    <xsd:sequence>
      <xsd:element name="bg"          type="CT_Background"       minOccurs="0"/>
      <xsd:element name="spTree"      type="CT_GroupShape"/>
      <xsd:element name="custDataLst" type="CT_CustomerDataList" minOccurs="0"/>
      <xsd:element name="controls"    type="CT_ControlList"      minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="name" type="xsd:string" default=""/>
  </xsd:complexType>

  <xsd:complexType name="CT_SlideTransition">
    <xsd:sequence>
      <xsd:choice minOccurs="0">
        <xsd:element name="blinds"    type="CT_OrientationTransition"/>
        <xsd:element name="checker"   type="CT_OrientationTransition"/>
        <xsd:element name="circle"    type="CT_Empty"/>
        <xsd:element name="dissolve"  type="CT_Empty"/>
        <xsd:element name="comb"      type="CT_OrientationTransition"/>
        <xsd:element name="cover"     type="CT_EightDirectionTransition"/>
        <xsd:element name="cut"       type="CT_OptionalBlackTransition"/>
        <xsd:element name="diamond"   type="CT_Empty"/>
        <xsd:element name="fade"      type="CT_OptionalBlackTransition"/>
        <xsd:element name="newsflash" type="CT_Empty"/>
        <xsd:element name="plus"      type="CT_Empty"/>
        <xsd:element name="pull"      type="CT_EightDirectionTransition"/>
        <xsd:element name="push"      type="CT_SideDirectionTransition"/>
        <xsd:element name="random"    type="CT_Empty"/>
        <xsd:element name="randomBar" type="CT_OrientationTransition"/>
        <xsd:element name="split"     type="CT_SplitTransition"/>
        <xsd:element name="strips"    type="CT_CornerDirectionTransition"/>
        <xsd:element name="wedge"     type="CT_Empty"/>
        <xsd:element name="wheel"     type="CT_WheelTransition"/>
        <xsd:element name="wipe"      type="CT_SideDirectionTransition"/>
        <xsd:element name="zoom"      type="CT_InOutTransition"/>
      </xsd:choice>
      <xsd:element name="sndAc"  type="CT_TransitionSoundAction" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionListModify"   minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="spd"      type="ST_TransitionSpeed" default="fast"/>
    <xsd:attribute name="advClick" type="xsd:boolean"        default="true"/>
    <xsd:attribute name="advTm"    type="xsd:unsignedInt"/>
  </xsd:complexType>

  <xsd:complexType name="CT_SlideTiming">
    <xsd:sequence>
      <xsd:element name="tnLst"  type="CT_TimeNodeList"        minOccurs="0"/>
      <xsd:element name="bldLst" type="CT_BuildList"           minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>
