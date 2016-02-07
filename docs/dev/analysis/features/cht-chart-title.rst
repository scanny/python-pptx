
Chart Title
===========

A chart has a title which contains a text_frame that you can write
text to. However, before you can add a title, you must set has_title
to True.


MS API Protocol
---------------

::

    >>> chart = ActiveDocument.InlineShapes(1).Chart
    >>> chart.HasTitle
    False
    >>> chart.HasTitle = True
    >>> chart.HasTitle
    True
    >>> chart.ChartTitle
    <ChartTitle object ...>
    >>> chart_title = chart.ChartTitle
    >>> chart_title.Text
    ''
    >>> chart_title.Text = "First Quarter Sales"
    >>> chart_title.Text
    'First Quarter Sales'


Python Usage
------------

::

    >>> chart.has_title
    False
    >>> chart.title.text_frame.text = 'Chart Title'
    ...
    AttributeError: 'NoneType' object has no attribute 'text_frame'
    >>> chart.has_title = True
    >>> chart.title.text_frame.text = 'Chart Title'
    >>> chart.title.text_frame.text
    u'Chart Title'


The chart titles use a TextFrame. Only the "text" property has been
reviewed and known to work for the chart titles. Some of the other
properties may function as well but only the text property has been
confirmed to work.


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_Chart">
    <xsd:sequence>
      <xsd:element name="title"            type="CT_Title"         minOccurs="0"/>
      <xsd:element name="autoTitleDeleted" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="pivotFmts"        type="CT_PivotFmts"     minOccurs="0"/>
      <xsd:element name="view3D"           type="CT_View3D"        minOccurs="0"/>
      <xsd:element name="floor"            type="CT_Surface"       minOccurs="0"/>
      <xsd:element name="sideWall"         type="CT_Surface"       minOccurs="0"/>
      <xsd:element name="backWall"         type="CT_Surface"       minOccurs="0"/>
      <xsd:element name="plotArea"         type="CT_PlotArea"/>
      <xsd:element name="legend"           type="CT_Legend"        minOccurs="0"/>
      <xsd:element name="plotVisOnly"      type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="dispBlanksAs"     type="CT_DispBlanksAs"  minOccurs="0"/>
      <xsd:element name="showDLblsOverMax" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="extLst"           type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Title">
    <xsd:sequence>
      <xsd:element name="tx"      type="CT_Tx"                minOccurs="0"/>
      <xsd:element name="layout"  type="CT_Layout"            minOccurs="0"/>
      <xsd:element name="overlay" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="spPr"    type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"    type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_ExtensionList"     minOccurs="0"/>
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

Minimal working XML for a chart with a title::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:chart>
      <c:title>
        <c:tx>
          <c:rich>
            <a:bodyPr />
            <a:p>
              <a:r>
                <a:t>Chart Title</a:t>
              </a:r>
            </a:p>
          </c:rich>
        </c:tx>
      </c:title>
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
          <c:crossAx val="-2068027336" />
          <c:crosses val="autoZero" />
        </c:valAx>
      </c:plotArea>
    </c:chart>
  </c:chartSpace>
