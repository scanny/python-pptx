
Slide Master
============

A slide master acts as a "parent" for zero or more slide layouts, providing
an inheritance base for placeholders and other slide and shape properties.


Schema excerpt
--------------

.. highlight:: xml

::

  <xsd:element name="sldMaster" type="CT_SlideMaster"/>

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

  <xsd:complexType name="CT_ColorMapping">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
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

  <xsd:complexType name="CT_SlideLayoutIdList">
    <xsd:sequence>
      <xsd:element name="sldLayoutId" type="CT_SlideLayoutIdListEntry" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SlideLayoutIdListEntry">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="id"  type="ST_SlideLayoutId"/>
    <xsd:attribute ref="r:id" type="ST_RelationshipId" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_RelationshipId">
    <xsd:restriction base="xsd:string"/>
  </xsd:simpleType>
