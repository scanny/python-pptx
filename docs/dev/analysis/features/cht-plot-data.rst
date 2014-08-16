
Chart - Plot Data
=================

The values and categories of each plot is specified by a data set. In the
typical case, that data is composed of a sequence of categories and
a sequence of series, where each series has a numeric value corresponding to
each of the categories.


Candidate protocol
------------------

::

    >>> plot = chart.plots[0]
    >>> plot.categories
    ('Foo', 'Bar', 'Baz')
    >>> len(plot.series)
    3
    >>> series = iter(plot.series)
    >>> next(series).values
    (1.2, 2.3, 3.4)
    >>> next(series).values
    (4.5, 5.6, 6.7)
    >>> next(series).values
    (7.8, 8.9, 9.0)


Feature Summary
---------------

* **Plot.categories** -- Read/only for now. Returns tuple of string.
* **Series.values** -- Read/only for now. Returns tuple of float.


Microsoft API
-------------

ChartGroup.CategoryCollection
    Returns all the visible categories in the chart group, or the specified
    visible category.

ChartGroup.SeriesCollection
    Returns all the series in the chart group.

Series.Values
    Returns or sets a collection of all the values in the series. Read/write
    Variant. Returns an array of float; accepts a spreadsheet formula or an
    array of numeric values.

Series.Formula
    Returns or sets the object's formula in A1-style notation. Read/write
    String.


XML specimens
-------------

.. highlight:: xml

Example series XML::

  <c:ser>
    <!-- ... -->
    <c:cat>
      <c:strRef>
        <c:f>Sheet1!$A$2:$A$4</c:f>
        <c:strCache>
          <c:ptCount val="3"/>
          <c:pt idx="0">
            <c:v>Foo</c:v>
          </c:pt>
          <c:pt idx="1">
            <c:v>Bar</c:v>
          </c:pt>
          <c:pt idx="2">
            <c:v>Baz</c:v>
          </c:pt>
        </c:strCache>
      </c:strRef>
    </c:cat>
    <c:val>
      <c:numRef>
        <c:f>Sheet1!$B$2:$B$4</c:f>
        <c:numCache>
          <c:ptCount val="3"/>
          <c:pt idx="0">
            <c:v>1.2</c:v>
          </c:pt>
          <c:pt idx="1">
            <c:v>2.3</c:v>
          </c:pt>
          <c:pt idx="2">
            <c:v>3.4</c:v>
          </c:pt>
        </c:numCache>
      </c:numRef>
    </c:val>
  </c:ser>


Related Schema Definitions
--------------------------

::

  <xsd:complexType name="CT_BarSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"              type="CT_UnsignedInt"/>
      <xsd:element name="order"            type="CT_UnsignedInt"/>
      <xsd:element name="tx"               type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"             type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="invertIfNegative" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="pictureOptions"   type="CT_PictureOptions"    minOccurs="0"/>
      <xsd:element name="dPt"              type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"            type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="trendline"        type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"          type="CT_ErrBars"           minOccurs="0"/>
      <xsd:element name="cat"              type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"              type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="shape"            type="CT_Shape"             minOccurs="0"/>
      <xsd:element name="extLst"           type="CT_ExtensionList"     minOccurs="0"/>
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

  <xsd:complexType name="CT_StrVal">
    <xsd:sequence>
      <xsd:element name="v" type="s:ST_Xstring"/>
    </xsd:sequence>
    <xsd:attribute name="idx" type="xsd:unsignedInt" use="required"/>
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

  <xsd:complexType name="CT_NumVal">
    <xsd:sequence>
      <xsd:element name="v" type="s:ST_Xstring"/>
    </xsd:sequence>
    <xsd:attribute name="idx"        type="xsd:unsignedInt" use="required"/>
    <xsd:attribute name="formatCode" type="s:ST_Xstring"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Xstring">
    <xsd:restriction base="xsd:string"/>
  </xsd:simpleType>
