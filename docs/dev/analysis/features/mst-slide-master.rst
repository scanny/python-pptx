
Slide Master
============

A slide master acts as a "parent" for zero or more slide layouts, providing
an inheritance base for placeholders and other slide and shape properties.


Schema excerpt
--------------

.. highlight:: xml

::

  <xsd:element name="sldMaster" type="CT_SlideMaster"/>

  <xsd:complexType name="CT_SlideMaster">
    <xsd:sequence>
      <xsd:element name="cSld" type="CT_CommonSlideData"/>
      <xsd:group   ref="EG_TopLevelSlide"/>
      <xsd:element name="sldLayoutIdLst" type="CT_SlideLayoutIdList"     minOccurs="0"/>
      <xsd:element name="transition"     type="CT_SlideTransition"       minOccurs="0"/>
      <xsd:element name="timing"         type="CT_SlideTiming"           minOccurs="0"/>
      <xsd:element name="hf"             type="CT_HeaderFooter"          minOccurs="0"/>
      <xsd:element name="txStyles"       type="CT_SlideMasterTextStyles" minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionListModify"   minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="preserve" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_SlideLayoutIdList">
    <xsd:sequence>
      <xsd:element name="sldLayoutId" type="CT_SlideLayoutIdListEntry"
                   minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SlideLayoutIdListEntry">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="id" type="ST_SlideLayoutId" use="optional"/>
    <xsd:attribute ref="r:id"                        use="required"/>
  </xsd:complexType>
