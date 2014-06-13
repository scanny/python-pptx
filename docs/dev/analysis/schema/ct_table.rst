############
``CT_Table``
############

.. highlight:: xml

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 15, 50

   Schema Name  , CT_Table
   Spec Name    , Table
   Tag(s)       , a:tbl
   Namespace    , drawingml (dml-main.xsd)
   Schema Line  , 2423
   Spec Section , 21.1.3.13


Resources
=========

* ISO-IEC-29500-1, Section 21.1.3 (DrawingML) Tables, pp3331
* ISO-IEC-29500-1, Section 21.1.3.13 tbl (Table), pp3344


Spec text
=========

   This element is the root element for a table. Within this element is
   contained everything that one would need to define a table within DrawingML.


Schema excerpt
==============

::

  <xsd:element name="tbl" type="CT_Table"/>

  <xsd:complexType name="CT_Table">
    <xsd:sequence>
      <xsd:element name="tblPr"   type="CT_TableProperties" minOccurs="0"/>
      <xsd:element name="tblGrid" type="CT_TableGrid"/>
      <xsd:element name="tr"      type="CT_TableRow"        minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TableProperties">
    <xsd:sequence>
      <xsd:group   ref="EG_FillProperties"   minOccurs="0"/>
      <xsd:group   ref="EG_EffectProperties" minOccurs="0"/>
      <xsd:choice minOccurs="0">
        <xsd:element name="tableStyle"   type="CT_TableStyle"/>
        <xsd:element name="tableStyleId" type="s:ST_Guid"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="rtl"      type="xsd:boolean" default="false"/>
    <xsd:attribute name="firstRow" type="xsd:boolean" default="false"/>
    <xsd:attribute name="firstCol" type="xsd:boolean" default="false"/>
    <xsd:attribute name="lastRow"  type="xsd:boolean" default="false"/>
    <xsd:attribute name="lastCol"  type="xsd:boolean" default="false"/>
    <xsd:attribute name="bandRow"  type="xsd:boolean" default="false"/>
    <xsd:attribute name="bandCol"  type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TableGrid">
    <xsd:sequence>
      <xsd:element name="gridCol" type="CT_TableCol" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TableCol">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="w" type="ST_Coordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TableRow">
    <xsd:sequence>
      <xsd:element name="tc"     type="CT_TableCell"              minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="h" type="ST_Coordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TableCell">
    <xsd:sequence>
      <xsd:element name="txBody" type="CT_TextBody"               minOccurs="0"/>
      <xsd:element name="tcPr"   type="CT_TableCellProperties"    minOccurs="0"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="rowSpan"  type="xsd:int"     default="1"/>
    <xsd:attribute name="gridSpan" type="xsd:int"     default="1"/>
    <xsd:attribute name="hMerge"   type="xsd:boolean" default="false"/>
    <xsd:attribute name="vMerge"   type="xsd:boolean" default="false"/>
    <xsd:attribute name="id"       type="xsd:string"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TextBody">
    <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle" minOccurs="0"/>
      <xsd:element name="p"        type="CT_TextParagraph" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TableCellProperties">
    <xsd:sequence>
      <xsd:element name="lnL"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="lnR"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="lnT"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="lnB"      type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="lnTlToBr" type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="lnBlToTr" type="CT_LineProperties"         minOccurs="0"/>
      <xsd:element name="cell3D"   type="CT_Cell3D"                 minOccurs="0"/>
      <xsd:group   ref="EG_FillProperties"                          minOccurs="0"/>
      <xsd:element name="headers"  type="CT_Headers"                minOccurs="0"/>
      <xsd:element name="extLst"   type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="marL"         type="ST_Coordinate32"         default="91440"/>
    <xsd:attribute name="marR"         type="ST_Coordinate32"         default="91440"/>
    <xsd:attribute name="marT"         type="ST_Coordinate32"         default="45720"/>
    <xsd:attribute name="marB"         type="ST_Coordinate32"         default="45720"/>
    <xsd:attribute name="vert"         type="ST_TextVerticalType"     default="horz"/>
    <xsd:attribute name="anchor"       type="ST_TextAnchoringType"    default="t"/>
    <xsd:attribute name="anchorCtr"    type="xsd:boolean"             default="false"/>
    <xsd:attribute name="horzOverflow" type="ST_TextHorzOverflowType" default="clip"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Coordinate">
    <xsd:union memberTypes="ST_CoordinateUnqualified s:ST_UniversalMeasure"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_CoordinateUnqualified">
    <xsd:restriction base="xsd:long">
      <xsd:minInclusive value="-27273042329600"/>
      <xsd:maxInclusive value="27273042316900"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_UniversalMeasure">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)"/>
    </xsd:restriction>
  </xsd:simpleType>
