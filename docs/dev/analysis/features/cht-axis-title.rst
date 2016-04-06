
Axis Title
==========

A chart axis has a title which contains a text_frame that you can write
text to. However, before you can add a title, you must set has_title
to True.


MS API Protocol
---------------

::

    >>> chart = ActiveDocument.InlineShapes(1).Chart
    >>> category_axis = chart.Axes(XlAxisType.xlCategory)
    >>> category_axis.HasTitle
    False
    >>> category_axis.HasTitle = True
    >>> category_axis.HasTitle
    True
    >>> category_axis.AxisTitle
    <AxisTitle object ...>
    >>> category_axis.AxisTitle.Text
    ''
    >>> category_axis.AxisTitle.Text = "Category Axis Title"
    >>> category_axis.AxisTitle.Text
    'Category Axis Title'


    >>> chart = ActiveDocument.InlineShapes(1).Chart
    >>> value_axis = chart.Axes(XlAxisType.xlValue)
    >>> value_axis.HasTitle
    False
    >>> value_axis.HasTitle = True
    >>> value_axis.HasTitle
    True
    >>> value_axis.AxisTitle
    <AxisTitle object ...>
    >>> value_axis.AxisTitle.Text
    ''
    >>> value_axis.AxisTitle.Text = "Value Axis Title"
    >>> value_axis.AxisTitle.Text
    'Value Axis Title'


Python Usage
------------

::

    >>> chart.category_axis.has_title
    False
    >>> chart.category_axis.title.text_frame.text = 'Category Axis Title'
    ...
    AttributeError: 'NoneType' object has no attribute 'text_frame'
    >>> chart.category_axis.has_title = True
    >>> chart.category_axis.title.text_frame.text = 'Category Axis Title'
    >>> chart.category_axis.title.text_frame.text
    u'Category Axis Title'


    >>> chart.value_axis.has_title
    False
    >>> chart.value_axis.title.text_frame.text = 'Value Axis Title'
    ...
    AttributeError: 'NoneType' object has no attribute 'text_frame'
    >>> chart.value_axis.has_title = True
    >>> chart.value_axis.title.text_frame.text = 'Value Axis Title'
    >>> chart.value_axis.title.text_frame.text
    u'Value Axis Title'


The axis titles use a TextFrame. Only the "text" property has been
reviewed and known to work for the axis titles. Some of the other
properties may function as well but only the text property has been
confirmed to work.


Related Schema Definitions
--------------------------

.. highlight:: xml

::

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

  <xsd:complexType name="CT_Title">
    <xsd:sequence>
      <xsd:element name="tx" type="CT_Tx" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="overlay" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Tx">
    <xsd:sequence>
      <xsd:choice>
        <xsd:element name="strRef" type="CT_StrRef"/>
        <xsd:element name="rich"   type="a:CT_TextBody"/>
      </xsd:choice>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Layout">
    <xsd:sequence>
      <xsd:element name="manualLayout" type="CT_ManualLayout"  minOccurs="0"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"               type="CT_Transform2D"            minOccurs="0"/>
      <xsd:group   ref="EG_Geometry"                                          minOccurs="0"/>
      <xsd:group   ref="EG_FillProperties"                                    minOccurs="0"/>
      <xsd:element name="ln"                 type="CT_LineProperties"         minOccurs="0"/>
      <xsd:group   ref="EG_EffectProperties"                                  minOccurs="0"/>
      <xsd:element name="scene3d"            type="CT_Scene3D"                minOccurs="0"/>
      <xsd:element name="sp3d"               type="CT_Shape3D"                minOccurs="0"/>
      <xsd:element name="extLst"             type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode"             type="ST_BlackWhiteMode"         use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TextBody">
    <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle"      minOccurs="0"/>
      <xsd:element name="p"        type="CT_TextParagraph"      maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>


XML specimens
-------------

.. highlight:: xml

Minimal working XML for a chart with axis titles::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:chart>
      <c:plotArea>
        <c:barChart>
          <c:barDir val="col" />
          <c:grouping val="clustered" />
          <c:axId val="-2068027336" />
          <c:axId val="-2113994440" />
        </c:barChart>
        <c:catAx>
          <c:axId val="-2068027336" />
          <c:scaling />
          <c:delete val="0" />
          <c:axPos val="b" />
          <c:title>
            <c:tx>
              <c:rich>
                <a:bodyPr />
                <a:lstStyle />
                <a:p>
                  <a:pPr>
                    <a:defRPr />
                  </a:pPr>
                  <a:r>
                    <a:t>Category Axis Title</a:t>
                  </a:r>
                </a:p>
              </c:rich>
            </c:tx>
          </c:title>
          <c:crossAx val="-2113994440" />
          <c:crosses val="autoZero" />
          <c:lblAlgn val="ctr" />
          <c:lblOffset val="100" />
        </c:catAx>
        <c:valAx>
          <c:axId val="-2113994440" />
          <c:scaling />
          <c:delete val="0" />
          <c:axPos val="l" />
          <c:title>
            <c:tx>
              <c:rich>
                <a:bodyPr />
                <a:lstStyle />
                <a:p>
                  <a:pPr>
                    <a:defRPr />
                  </a:pPr>
                  <a:r>
                    <a:t>Value Axis Title</a:t>
                  </a:r>
                </a:p>
              </c:rich>
            </c:tx>
          </c:title>
          <c:crossAx val="-2068027336" />
          <c:crosses val="autoZero" />
        </c:valAx>
      </c:plotArea>
    </c:chart>
  </c:chartSpace>
