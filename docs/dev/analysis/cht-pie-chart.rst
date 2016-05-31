
Pie Chart
=========

A pie chart is one of the fundamental plot types.


XML specimens
-------------

.. highlight:: xml

Minimal working XML for a pie plot. Note this does not reference values in a
spreadsheet.::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:chart>
      <c:plotArea>
        <c:pieChart>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx>
              <c:v>Sales</c:v>
            </c:tx>
            <c:cat>
              <c:strLit>
                <c:ptCount val="4"/>
                <c:pt idx="0">
                  <c:v>1st Qtr</c:v>
                </c:pt>
                <c:pt idx="1">
                  <c:v>2nd Qtr</c:v>
                </c:pt>
                <c:pt idx="2">
                  <c:v>3rd Qtr</c:v>
                </c:pt>
                <c:pt idx="3">
                  <c:v>4th Qtr</c:v>
                </c:pt>
              </c:strLit>
            </c:cat>
            <c:val>
              <c:numLit>
                <c:ptCount val="4"/>
                <c:pt idx="0">
                  <c:v>8.2</c:v>
                </c:pt>
                <c:pt idx="1">
                  <c:v>3.2</c:v>
                </c:pt>
                <c:pt idx="2">
                  <c:v>1.4</c:v>
                </c:pt>
                <c:pt idx="3">
                  <c:v>1.2</c:v>
                </c:pt>
              </c:numLit>
            </c:val>
          </c:ser>
        </c:pieChart>
      </c:plotArea>
    </c:chart>
  </c:chartSpace>


Related Schema Definitions
--------------------------

.. highlight:: xml

Pie chart elements::

  <xsd:complexType name="CT_PieChart">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="varyColors"    type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="ser"           type="CT_PieSer"        minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"         type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="firstSliceAng" type="CT_FirstSliceAng" minOccurs="0"/>
      <xsd:element name="extLst"        type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_PieSer">
    <xsd:sequence>
      <xsd:element name="idx"       type="CT_UnsignedInt"/>
      <xsd:element name="order"     type="CT_UnsignedInt"/>
      <xsd:element name="tx"        type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"      type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="explosion" type="CT_UnsignedInt"       minOccurs="0"/>
      <xsd:element name="dPt"       type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"     type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="cat"       type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"       type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="extLst"    type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SerTx">
    <xsd:sequence>
      <xsd:choice minOccurs="1" maxOccurs="1">
        <xsd:element name="strRef" type="CT_StrRef"/>
        <xsd:element name="v"      type="s:ST_Xstring"/>
      </xsd:choice>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_AxDataSource">
    <xsd:sequence>
      <xsd:choice minOccurs="1" maxOccurs="1">
        <xsd:element name="multiLvlStrRef" type="CT_MultiLvlStrRef"/>
        <xsd:element name="numRef"         type="CT_NumRef"/>
        <xsd:element name="numLit"         type="CT_NumData"/>
        <xsd:element name="strRef"         type="CT_StrRef"/>
        <xsd:element name="strLit"         type="CT_StrData"/>
      </xsd:choice>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_NumDataSource">
    <xsd:sequence>
      <xsd:choice minOccurs="1" maxOccurs="1">
        <xsd:element name="numRef" type="CT_NumRef"/>
        <xsd:element name="numLit" type="CT_NumData"/>
      </xsd:choice>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_NumRef">
    <xsd:sequence>
      <xsd:element name="f"        type="xsd:string"/>
      <xsd:element name="numCache" type="CT_NumData"       minOccurs="0"/>
      <xsd:element name="extLst"   type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_NumData">
    <xsd:sequence>
      <xsd:element name="formatCode" type="s:ST_Xstring"     minOccurs="0"/>
      <xsd:element name="ptCount"    type="CT_UnsignedInt"   minOccurs="0"/>
      <xsd:element name="pt"         type="CT_NumVal"        minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_StrRef">
    <xsd:sequence>
      <xsd:element name="f"        type="xsd:string"/>
      <xsd:element name="strCache" type="CT_StrData"       minOccurs="0"/>
      <xsd:element name="extLst"   type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_StrData">
    <xsd:sequence>
      <xsd:element name="ptCount" type="CT_UnsignedInt"   minOccurs="0"/>
      <xsd:element name="pt"      type="CT_StrVal"        minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="extLst"  type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:simpleType name="ST_Xstring">
    <xsd:restriction base="xsd:string"/>
  </xsd:simpleType>

  <xsd:complexType name="CT_NumVal">
    <xsd:sequence>
      <xsd:element name="v" type="s:ST_Xstring"/>
    </xsd:sequence>
    <xsd:attribute name="idx"        type="xsd:unsignedInt" use="required"/>
    <xsd:attribute name="formatCode" type="s:ST_Xstring"/>
  </xsd:complexType>
