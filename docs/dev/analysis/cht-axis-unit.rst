
Chart - ValueAxis.major/minor_unit
==================================

The value axis has major and minor divisions, corresponding to where tick
marks, tick labels, and gridlines appear, when present. How frequently these
appear is determined by the major/minor units setting, specified as
a floating point number of units to skip between divisions. These settings
may be specified explictly, or PowerPoint can determine a sensible default
based on the chart data. The latter setting is labeled 'Auto' in the UI. By
default, major and minor unit are both set to 'Auto' on new charts.


Candidate protocol
------------------

The properties ``ValueAxis.major_unit`` and ``ValueAxis.minor_unit`` are used
to access and change this setting.

|None| is used as an out-of-band value to signify `Auto` behavior. No
separate boolean properties are required.

::

    >>> value_axis = chart.value_axis
    >>> value_axis.major_unit
    None
    >>> value_axis.major_unit = 10
    >>> value_axis.major_unit
    10.0
    >>> value_axis.major_unit = -4.2
    Traceback ...
    ValueError: must be positive numeric value
    >>> value_axis.major_unit = None
    >>> value_axis.major_unit
    None


Microsoft API
-------------

Axis.MajorUnit
    Returns or sets the major units for the value axis. Read/write Double.

Axis.MinorUnit
    Returns or sets the minor units on the value axis. Read/write Double.

Axis.MajorUnitIsAuto
    True if PowerPoint calculates the major units for the value axis.
    Read/write Boolean.

Axis.MinorUnitIsAuto
    True if PowerPoint calculates minor units for the value axis. Read/write
    Boolean.


PowerPoint behavior
-------------------

Major and minor unit values are viewed and changed using the `Scale` pane of
the `Format Axis` dialog. Checkboxes are used to set a value to `Auto`.
Changing the floating point value causes the `Auto` checkbox to turn off.


XML Semantics
-------------

* Only a value axis axis or date axis can have a ``<c:majorUnit>`` or
  a ``<c:minorUnit>`` element.
* `Auto` behavior is signified by having no element for that unit.


XML specimens
-------------

.. highlight:: xml

Example value axis XML having an override for major unit::

  <c:valAx>
    <c:axId val="-2101345848"/>
    <c:scaling>
      <c:orientation val="minMax"/>
    </c:scaling>
    <c:delete val="0"/>
    <c:axPos val="l"/>
    <c:majorGridlines/>
    <c:numFmt formatCode="General" sourceLinked="1"/>
    <c:majorTickMark val="out"/>
    <c:minorTickMark val="none"/>
    <c:tickLblPos val="nextTo"/>
    <c:crossAx val="-2030568888"/>
    <c:crosses val="autoZero"/>
    <c:crossBetween val="between"/>
    <c:majorUnit val="20.0"/>
  </c:valAx>


Related Schema Definitions
--------------------------

::

  <xsd:complexType name="CT_ValAx">
    <xsd:sequence>
      <xsd:group ref="EG_AxShared" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="crossBetween" type="CT_CrossBetween"  minOccurs="0"/>
      <xsd:element name="majorUnit"    type="CT_AxisUnit"      minOccurs="0"/>
      <xsd:element name="minorUnit"    type="CT_AxisUnit"      minOccurs="0"/>
      <xsd:element name="dispUnits"    type="CT_DispUnits"     minOccurs="0"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_DateAx">
    <xsd:sequence>
      <xsd:group    ref="EG_AxShared"/>
      <xsd:element name="auto"          type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="lblOffset"     type="CT_LblOffset"     minOccurs="0"/>
      <xsd:element name="baseTimeUnit"  type="CT_TimeUnit"      minOccurs="0"/>
      <xsd:element name="majorUnit"     type="CT_AxisUnit"      minOccurs="0"/>
      <xsd:element name="majorTimeUnit" type="CT_TimeUnit"      minOccurs="0"/>
      <xsd:element name="minorUnit"     type="CT_AxisUnit"      minOccurs="0"/>
      <xsd:element name="minorTimeUnit" type="CT_TimeUnit"      minOccurs="0"/>
      <xsd:element name="extLst"        type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="EG_AxShared">
    <xsd:sequence>
      <xsd:element name="axId"           type="CT_UnsignedInt"/>
      <xsd:element name="scaling"        type="CT_Scaling"/>
      <xsd:element name="delete"         type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="axPos"          type="CT_AxPos"/>
      <xsd:element name="majorGridlines" type="CT_ChartLines"        minOccurs="0"/>
      <xsd:element name="minorGridlines" type="CT_ChartLines"        minOccurs="0"/>
      <xsd:element name="title"          type="CT_Title"             minOccurs="0"/>
      <xsd:element name="numFmt"         type="CT_NumFmt"            minOccurs="0"/>
      <xsd:element name="majorTickMark"  type="CT_TickMark"          minOccurs="0"/>
      <xsd:element name="minorTickMark"  type="CT_TickMark"          minOccurs="0"/>
      <xsd:element name="tickLblPos"     type="CT_TickLblPos"        minOccurs="0"/>
      <xsd:element name="spPr"           type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"           type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="crossAx"        type="CT_UnsignedInt"/>
      <xsd:choice minOccurs="0" maxOccurs="1">
        <xsd:element name="crosses"   type="CT_Crosses"/>
        <xsd:element name="crossesAt" type="CT_Double"/>
      </xsd:choice>
    </xsd:sequence>
  </xsd:group>

  <xsd:complexType name="CT_AxisUnit">
    <xsd:attribute name="val" type="ST_AxisUnit" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_AxisUnit">
    <xsd:restriction base="xsd:double">
      <xsd:minExclusive value="0"/>
    </xsd:restriction>
  </xsd:simpleType>
