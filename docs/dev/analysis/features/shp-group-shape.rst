
Group Shape
===========

Group of shapes that appear as a single shape ...


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_GroupShape">
    <xsd:sequence>
      <xsd:element name="nvGrpSpPr" type="CT_GroupShapeNonVisual"    minOccurs="1" maxOccurs="1"/>
      <xsd:element name="grpSpPr"   type="a:CT_GroupShapeProperties" minOccurs="1" maxOccurs="1"/>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="sp"           type="CT_Shape"/>
        <xsd:element name="grpSp"        type="CT_GroupShape"/>
        <xsd:element name="graphicFrame" type="CT_GraphicalObjectFrame"/>
        <xsd:element name="cxnSp"        type="CT_Connector"/>
        <xsd:element name="pic"          type="CT_Picture"/>
        <xsd:element name="contentPart"  type="CT_Rel"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GroupShapeNonVisual">
    <xsd:sequence>
      <xsd:element name="cNvPr"      type="a:CT_NonVisualDrawingProps"           minOccurs="1" maxOccurs="1"/>
      <xsd:element name="cNvGrpSpPr" type="a:CT_NonVisualGroupDrawingShapeProps" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="nvPr"       type="CT_ApplicationNonVisualDrawingProps"  minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GroupShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"    type="CT_GroupTransform2D"       minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_FillProperties"                         minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_EffectProperties"                       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="scene3d" type="CT_Scene3D"                minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
  </xsd:complexType>
