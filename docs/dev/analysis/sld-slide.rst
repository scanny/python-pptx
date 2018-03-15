.. _Slide:

Slide
=====

A slide is the fundamental visual content container in a presentation, that
content taking the form of shape objects.

The slides in a presentation are owned by the presentation object. In
particular, the unique identifier of a slide, the slide id, is assigned and
managed by the presentation part, and is not recorded in the slide XML.

A slide master and slide layout are both closely related to slide and the three
share the majority of their properties and behaviors.


Slide ID
--------

PowerPoint assigns a unique integer identifier to a slide when it is created.
Note that this identifier is only present in the presentation part, and maps
to a relationship ID; it is not recorded in the XML of the slide itself.
Therefore the identifier is only unique within a presentation. The ID takes
the form of an integer starting at 256 and incrementing by one for each new
slide. Changing the ordering of the slides does not change the id. The id of
a deleted slide is not reused, although I'm not sure whether it's clever
enough not to reuse the id of the last added slide when it's been deleted as
there doesn't seem to be any record in the XML of the max value assigned.


XML specimens
-------------

.. highlight:: xml


Example presentation XML showing sldIdLst::

  <p:presentation
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
    <p:sldMasterIdLst>
      <p:sldMasterId id="2147483648" r:id="rId1"/>
    </p:sldMasterIdLst>

    <p:sldIdLst>
      <p:sldId r:id="rId7" id="256"/>
    </p:sldIdLst>

    <p:sldSz cx="12192000" cy="6858000"/>
    <p:notesSz cx="6858000" cy="9144000"/>
  </p:presentation>

Example slide contents::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <p:sld
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
    <p:cSld>
      <p:spTree>
        <p:nvGrpSpPr>
          <p:cNvPr id="1" name=""/>
          <p:cNvGrpSpPr/>
          <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr>
            <p:cNvPr id="2" name="Chart 1"/>
            <p:cNvGraphicFramePr>
              <a:graphicFrameLocks noGrp="1"/>
            </p:cNvGraphicFramePr>
            <p:nvPr/>
          </p:nvGraphicFramePr>
          <p:xfrm>
            <a:off x="3801533" y="956733"/>
            <a:ext cx="8128000" cy="5418667"/>
          </p:xfrm>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>
            </a:graphicData>
          </a:graphic>
        </p:graphicFrame>
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

  <xsd:complexType name="CT_Presentation">
    <!-- ... -->
    <xsd:element name="sldIdLst" type="CT_SlideIdList" minOccurs="0"/>
    <!-- ... -->
  </xsd:complexType>

  <xsd:complexType name="CT_SlideIdList">
    <xsd:sequence>
      <xsd:element name="sldId" type="CT_SlideIdListEntry" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SlideIdListEntry">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="id" type="ST_SlideId" use="required"/>
    <xsd:attribute ref="r:id" use="required"/>
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

  <xsd:complexType name="CT_Background">
    <xsd:choice>
      <xsd:element name="bgPr"  type="CT_BackgroundProperties"/>
      <xsd:element name="bgRef" type="a:CT_StyleMatrixReference"/>
    </xsd:choice>
    <xsd:attribute name="bwMode" type="a:ST_BlackWhiteMode" use="optional" default="white"/>
  </xsd:complexType>

  <xsd:complexType name="CT_BackgroundProperties">
    <xsd:sequence>
      <xsd:group ref="a:EG_FillProperties"/>
      <xsd:group ref="a:EG_EffectProperties" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="shadeToTitle" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_SlideId">
    <xsd:restriction base="xsd:unsignedInt">
      <xsd:minInclusive value="256"/>
      <xsd:maxExclusive value="2147483648"/>
    </xsd:restriction>
  </xsd:simpleType>
