.. _cht-date-axis:

Chart - Date Axis
=================

A date axis maintains a linear relationship between axis distance and elapsed
time in days, months, or years, even if the data items do not include values
for each base time unit. This is in contrast to a category axis, which uses
the independent axis values themselves to determine the discrete set of
categories.

The date axis does this by essentially creating a "logical category" for each
base time unit in a range, whether the data contains a value for that time
unit or not.

Variations in the number of days in a month (or year) is also taken care of
automatically. This would otherwise cause month starts to appear at days not
the first of the month.

This does not make the category axis continuous, however the "grain" of the
discrete units can be quite small in comparison to the axis length. The
individual base unit divisions can but are not typically represented
graphically. The granularity of major and minor divisions is specified as
a property on the date axis.


MS API
------

* Axis.CategoryType = XL_CATEGORY_TYPE.TIME
* Axis.BaseUnit = XL_TIME_UNIT.DAYS


Enumerations
------------

XL_CATEGORY_TYPE
    * .AUTOMATIC (-4105)
    * .CATEGORY (2)
    * .TIME (3)

XL_TIME_UNIT
    https://msdn.microsoft.com/en-us/library/office/ff746136.aspx
    * .DAYS (0)
    * .MONTHS (1)
    * .YEARS (2)


XML Semantics
-------------

Changing category axis between `c:catAx` and `c:dateAx`
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

In the MS API, DateAxis.CategoryType is used to discover the category type,
but is also used to change an axis from a date axis to a category axis and
back. Note the corresponding property in |pp| is read-only for now.

Changing from a category to date axis leaves the base, major, and minor unit
elements in place. They apply again (instead of the defaults) if the axis is
changed back::

    <c:baseTimeUnit val="days"/>
    <c:majorUnit val="1.0"/>
    <c:majorTimeUnit val="months"/>
    <c:minorUnit val="1.0"/>


Choice of date axis
~~~~~~~~~~~~~~~~~~~

PowerPoint automatically uses a `c:dateAx` element if the category labels in
Excel are dates (numbers with date formatting).

|pp| uses a `c:dateAx` element if the category labels in the chart data
object are datetime.date or datetime.datetime objects.

Note that |pp| does not change the category axis type when using
`Chart.replace_data()`.

A date axis is only available on an area, bar (including column), or line
chart.


Multi-level categories and date axis
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Multi-level categories are mutually exclusive with a date axis.

A `c:multiLvlStrRef` element is the only one that can enclose a multi-level
category hierarchy. There is no provision for any or all the levels under
this element to be other than string values (as the 'Str' in its name
implies).


Automatic axis type
~~~~~~~~~~~~~~~~~~~

The `c:auto` element under `c:catAx` or `c:dateAx` controls whether the
category axis changes automatically between category and date types.


Specimen XML
------------

.. highlight:: xml

Series having date categories::

  <c:ser>
    <c:idx val="0"/>
    <c:order val="0"/>
    <c:tx>
      <c:strRef>
        <c:f>Sheet1!$B$1</c:f>
        <c:strCache>
          <c:ptCount val="1"/>
          <c:pt idx="0">
            <c:v>Buzz</c:v>
          </c:pt>
        </c:strCache>
      </c:strRef>
    </c:tx>
    <c:marker>
      <c:symbol val="none"/>
    </c:marker>
    <c:cat>
      <c:numRef>
        <c:f>Sheet1!$A$2:$A$520</c:f>
        <c:numCache>
          <c:formatCode>mm\-dd\-yyyy</c:formatCode>
          <c:ptCount val="3"/>
          <c:pt idx="0">
            <c:v>42156.0</c:v>
          </c:pt>
          <c:pt idx="1">
            <c:v>42157.0</c:v>
          </c:pt>
          <c:pt idx="2">
            <c:v>42158.0</c:v>
          </c:pt>
        </c:numCache>
      </c:numRef>
    </c:cat>
    <c:val>
      <c:numRef>
        <c:f>Sheet1!$B$2:$B$520</c:f>
        <c:numCache>
          <c:formatCode>0.0</c:formatCode>
          <c:ptCount val="3"/>
          <c:pt idx="0">
            <c:v>19.65943065775559</c:v>
          </c:pt>
          <c:pt idx="1">
            <c:v>20.13705095574664</c:v>
          </c:pt>
          <c:pt idx="2">
            <c:v>19.48264757927654</c:v>
          </c:pt>
        </c:numCache>
      </c:numRef>
    </c:val>

Plot area having date axis::

    <c:plotArea>
      <c:lineChart>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0">
                  <c:v>Series 1</c:v>
                </c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:numRef>
              <c:f>Sheet1!$A$2:$A$4</c:f>
              <c:numCache>
                <c:formatCode>d\-mmm</c:formatCode>
                <c:ptCount val="3"/>
                <c:pt idx="0">
                  <c:v>42370.0</c:v>
                </c:pt>
                <c:pt idx="1">
                  <c:v>42371.0</c:v>
                </c:pt>
                <c:pt idx="2">
                  <c:v>42372.0</c:v>
                </c:pt>
              </c:numCache>
            </c:numRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$4</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="3"/>
                <c:pt idx="0">
                  <c:v>4.3</c:v>
                </c:pt>
                <c:pt idx="1">
                  <c:v>2.5</c:v>
                </c:pt>
                <c:pt idx="2">
                  <c:v>3.5</c:v>
                </c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
          <c:smooth val="0"/>
        </c:ser>
        <c:axId val="2142588392"/>
        <c:axId val="2106388088"/>
      </c:lineChart>
      <c:dateAx>
        <c:axId val="2142588392"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="d\-mmm" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="2106388088"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblOffset val="100"/>
        <c:baseTimeUnit val="days"/>
      </c:dateAx>
      <c:valAx>
        <c:axId val="2106388088"/>
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
        <c:crossAx val="2142588392"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>


References
----------

* Understanding Date-Based Axis Versus Category-Based Axis in Trend Charts
  http://www.quepublishing.com/articles/article.aspx?p=1642672&seqNum=2


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_PlotArea">
    <xsd:sequence>
      <!-- 17 others -->
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="valAx"  type="CT_ValAx"/>
        <xsd:element name="catAx"  type="CT_CatAx"/>
        <xsd:element name="dateAx" type="CT_DateAx"/>
        <xsd:element name="serAx"  type="CT_SerAx"/>
      </xsd:choice>
      <xsd:element name="dTable" type="CT_DTable"            minOccurs="0"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_CatAx">  <!-- denormalized -->
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
      <xsd:choice                                                    minOccurs="0">
        <xsd:element name="crosses"      type="CT_Crosses"/>
        <xsd:element name="crossesAt"    type="CT_Double"/>
      </xsd:choice>
      <xsd:element name="auto"           type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="lblAlgn"        type="CT_LblAlgn"           minOccurs="0"/>
      <xsd:element name="lblOffset"      type="CT_LblOffset"         minOccurs="0"/>
      <xsd:element name="tickLblSkip"    type="CT_Skip"              minOccurs="0"/>
      <xsd:element name="tickMarkSkip"   type="CT_Skip"              minOccurs="0"/>
      <xsd:element name="noMultiLvlLbl"  type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_DateAx">  <!-- denormalized -->
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
      <xsd:choice                                                    minOccurs="0">
        <xsd:element name="crosses"      type="CT_Crosses"/>
        <xsd:element name="crossesAt"    type="CT_Double"/>
      </xsd:choice>
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

  <xsd:complexType name="CT_ValAx">  <!-- denormalized -->
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
      <xsd:choice                                                    minOccurs="0">
        <xsd:element name="crosses"   type="CT_Crosses"/>
        <xsd:element name="crossesAt" type="CT_Double"/>
      </xsd:choice>
      <xsd:element name="crossBetween"   type="CT_CrossBetween"      minOccurs="0"/>
      <xsd:element name="majorUnit"      type="CT_AxisUnit"          minOccurs="0"/>
      <xsd:element name="minorUnit"      type="CT_AxisUnit"          minOccurs="0"/>
      <xsd:element name="dispUnits"      type="CT_DispUnits"         minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SerAx">
    <xsd:sequence>
      <xsd:group   ref="EG_AxShared"/>
      <xsd:element name="tickLblSkip"  type="CT_Skip"          minOccurs="0"/>
      <xsd:element name="tickMarkSkip" type="CT_Skip"          minOccurs="0"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
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

  <xsd:complexType name="CT_ChartLines">
    <xsd:sequence>
      <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Crosses">
    <xsd:attribute name="val" type="ST_Crosses" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Scaling">
    <xsd:sequence>
      <xsd:element name="logBase"     type="CT_LogBase"       minOccurs="0"/>
      <xsd:element name="orientation" type="CT_Orientation"   minOccurs="0"/>
      <xsd:element name="max"         type="CT_Double"        minOccurs="0"/>
      <xsd:element name="min"         type="CT_Double"        minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_NumFmt">
    <xsd:attribute name="formatCode"   type="xsd:string"  use="required"/>
    <xsd:attribute name="sourceLinked" type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TickLblPos">
    <xsd:attribute name="val" type="ST_TickLblPos" default="nextTo"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TickMark">
    <xsd:attribute name="val" type="ST_TickMark" default="cross"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TimeUnit">
    <xsd:attribute name="val" type="ST_TimeUnit" default="days"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Boolean">
    <xsd:attribute name="val" type="xsd:boolean" default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Double">
    <xsd:attribute name="val" type="xsd:double" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_AxisUnit">
    <xsd:restriction base="xsd:double">
      <xsd:minExclusive value="0"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Crosses">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="autoZero"/>
      <xsd:enumeration value="max"/>
      <xsd:enumeration value="min"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TickLblPos">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="high"/>
      <xsd:enumeration value="low"/>
      <xsd:enumeration value="nextTo"/>
      <xsd:enumeration value="none"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TickMark">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="cross"/>
      <xsd:enumeration value="in"/>
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="out"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_TimeUnit">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="days"/>
      <xsd:enumeration value="months"/>
      <xsd:enumeration value="years"/>
    </xsd:restriction>
  </xsd:simpleType>
