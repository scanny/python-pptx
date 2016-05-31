.. _ChartTitle:


Chart Title
===========

A chart can have a title. The title is a rich text container, and can contain
arbitrary text with arbitrary formatting (font, size, color, etc.). There is
little but one thing to distinquish a chart title from an independent text
box; its position is automatically adjusted by the chart to account for
resizing and movement.

A title is visible whenever present. The only way to "hide" it is to delete
it, along with its contents.

Although it will not yet be supported, the chart title can be specified in
the XML as a cell reference in the Excel worksheet. In general, any
constructive operations on the title will remove this.


Candidate protocol
------------------

Chart title access
~~~~~~~~~~~~~~~~~~

A newly created chart has no title. ``Chart.chart_title`` is a valid call in
this state, but returns |None|::

    >>> chart = shapes.add_chart(...).chart
    >>> chart.has_title
    False
    >>> chart.chart_title
    None

Assigning |True| to ``.has_title`` causes an empty title element to be added
along with its text frame elements (when not already present)::

    >>> chart.has_title = True
    >>> chart.has_title
    True
    >>> chart.chart_title
    <pptx.chart.ChartTitle object at 0x65432fd>

Assigning |False| to ``.has_title`` removes the title element from the XML
along with its contents::

    >>> chart.has_title = False
    >>> chart.has_title
    False
    >>> chart.chart_title
    None

Assigning |None| to ``Chart.chart_title`` has the same effect (not sure we'll
actually implement this as a priority)::

    >>> chart.chart_title = None
    >>> chart.chart_title
    None


ChartTitle.text_frame
~~~~~~~~~~~~~~~~~~~~~

The ``ChartTitle`` object can contain either a text frame or an Excel cell
reference (``<c:strRef>``). However, the only operation on the `c:strRef`
element the library will support (for now anyway) is to delete it when adding
a text frame.

``ChartTitle.has_text_frame`` is used to determine whether a text
frame is present. Assigning |True| to ``.has_text_frame`` causes any Excel
reference to be removed and an empty text frame to be inserted. If a text
frame is already present, no changes are made::

    >>> chart_title.has_text_frame
    False
    >>> chart_title.has_text_frame = True
    >>> chart_title.has_text_frame
    True

The text frame can be accessed using ``ChartTitle.text_frame``. This call is
always valid, and will create a text frame if one is not present::

    >>> chart_title.text_frame
    <pptx.text.text.TextFrame object at 0x65432fe>

As a shortcut, the plain text string of a title can be accessed or set using
``Chart.title``. If no text frame is present, |None| is returned from
``.title`` used as a RHS expression. It is essentially identical to
``Chart.chart_title.text_frame.text``::

    >>> chart_title.has_text_frame
    False
    >>> chart.title
    None
    >>> chart.title = 'Foobar'
    >>> chart.title
    'Foobar'


XML semantics
-------------

* ``c:autoTitleDeleted`` set True has no effect on the visibility of a default
  chart title (no actual text, 'placeholder' display: 'Chart Title'. It also
  seems to have no effect on an actual title, having a text frame and actual
  text.


XML specimens
-------------

.. highlight:: xml

Default when clicking *Chart Title > Title Above Chart* from ribbon (before
changing the text of the title)::

  <c:chart>
    <c:title>
      <c:layout/>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      ...
    </c:plotArea>
  </c:chart>

Text 'Foobar' typed into chart title just after adding it from ribbon::

  <c:title>
    <c:tx>
      <c:rich>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:pPr>
            <a:defRPr/>
          </a:pPr>
          <a:r>
            <a:rPr lang="en-US" dirty="0" smtClean="0"/>
            <a:t>Foobar</a:t>
          </a:r>
          <a:endParaRPr lang="en-US" dirty="0"/>
        </a:p>
      </c:rich>
    </c:tx>
    <c:layout/>
    <c:overlay val="0"/>
  </c:title>


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

  <xsd:complexType name="CT_TextBody">
    <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle"      minOccurs="0"/>
      <xsd:element name="p"        type="CT_TextParagraph"      maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>
