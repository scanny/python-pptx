
Color
=====

ColorFormat object ...


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:group name="EG_ColorChoice">
    <xsd:choice>
      <xsd:element name="scrgbClr"  type="CT_ScRgbColor"  minOccurs="1" maxOccurs="1"/>
      <xsd:element name="srgbClr"   type="CT_SRgbColor"   minOccurs="1" maxOccurs="1"/>
      <xsd:element name="hslClr"    type="CT_HslColor"    minOccurs="1" maxOccurs="1"/>
      <xsd:element name="sysClr"    type="CT_SystemColor" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="schemeClr" type="CT_SchemeColor" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="prstClr"   type="CT_PresetColor" minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_ScRgbColor">
    <xsd:sequence>
      <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="r" type="ST_Percentage" use="required"/>
    <xsd:attribute name="g" type="ST_Percentage" use="required"/>
    <xsd:attribute name="b" type="ST_Percentage" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_SRgbColor">
    <xsd:sequence>
      <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="val" type="s:ST_HexColorRGB" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_HslColor">
    <xsd:sequence>
      <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="hue" type="ST_PositiveFixedAngle" use="required"/>
    <xsd:attribute name="sat" type="ST_Percentage"         use="required"/>
    <xsd:attribute name="lum" type="ST_Percentage"         use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_SystemColor">
    <xsd:sequence>
      <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="val"     type="ST_SystemColorVal" use="required"/>
    <xsd:attribute name="lastClr" type="s:ST_HexColorRGB"  use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_SchemeColor">
    <xsd:sequence>
      <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="val" type="ST_SchemeColorVal" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PresetColor">
    <xsd:sequence>
      <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="val" type="ST_PresetColorVal" use="required"/>
  </xsd:complexType>

  <xsd:group name="EG_ColorTransform">
    <xsd:choice>
      <xsd:element name="tint"     type="CT_PositiveFixedPercentage" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="shade"    type="CT_PositiveFixedPercentage" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="comp"     type="CT_ComplementTransform"     minOccurs="1" maxOccurs="1"/>
      <xsd:element name="inv"      type="CT_InverseTransform"        minOccurs="1" maxOccurs="1"/>
      <xsd:element name="gray"     type="CT_GrayscaleTransform"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="alpha"    type="CT_PositiveFixedPercentage" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="alphaOff" type="CT_FixedPercentage"         minOccurs="1" maxOccurs="1"/>
      <xsd:element name="alphaMod" type="CT_PositivePercentage"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="hue"      type="CT_PositiveFixedAngle"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="hueOff"   type="CT_Angle"                   minOccurs="1" maxOccurs="1"/>
      <xsd:element name="hueMod"   type="CT_PositivePercentage"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="sat"      type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="satOff"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="satMod"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="lum"      type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="lumOff"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="lumMod"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="red"      type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="redOff"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="redMod"   type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="green"    type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="greenOff" type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="greenMod" type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="blue"     type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="blueOff"  type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="blueMod"  type="CT_Percentage"              minOccurs="1" maxOccurs="1"/>
      <xsd:element name="gamma"    type="CT_GammaTransform"          minOccurs="1" maxOccurs="1"/>
      <xsd:element name="invGamma" type="CT_InverseGammaTransform"   minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_Percentage">
    <xsd:attribute name="val" type="ST_Percentage" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Percentage">
    <xsd:union memberTypes="ST_PercentageDecimal s:ST_Percentage"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_PercentageDecimal">
    <xsd:restriction base="xsd:int"/>
  </xsd:simpleType>

  <xsd:simpleType name="s:ST_Percentage">  <!-- denormalized -->
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="-?[0-9]+(\.[0-9]+)?%"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_HexColorRGB">
    <xsd:restriction base="xsd:hexBinary">
      <xsd:length value="3" fixed="true"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_SchemeColorVal">
    <xsd:restriction base="xsd:token">
      <xsd:enumeration value="bg1"/>
      <xsd:enumeration value="tx1"/>
      <xsd:enumeration value="bg2"/>
      <xsd:enumeration value="tx2"/>
      <xsd:enumeration value="accent1"/>
      <xsd:enumeration value="accent2"/>
      <xsd:enumeration value="accent3"/>
      <xsd:enumeration value="accent4"/>
      <xsd:enumeration value="accent5"/>
      <xsd:enumeration value="accent6"/>
      <xsd:enumeration value="hlink"/>
      <xsd:enumeration value="folHlink"/>
      <xsd:enumeration value="phClr"/>
      <xsd:enumeration value="dk1"/>
      <xsd:enumeration value="lt1"/>
      <xsd:enumeration value="dk2"/>
      <xsd:enumeration value="lt2"/>
    </xsd:restriction>
  </xsd:simpleType>
