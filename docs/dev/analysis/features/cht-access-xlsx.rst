
Chart - Embedded Worksheet
==========================

The data for a chart in PowerPoint is stored in an embedded Excel
spreadsheet.


Spike approach
--------------

* [ ] driver with one blank slide, builder column chart single series.
* [ ] Can edit Excel data in resulting file
* [ ] resulting file has an embedded Excel package of expected name
* [ ] XML uses worksheet references, not just cached data


Proposed protocol
-----------------

::

    >>> graphic_frame = prs.slides[0].shapes[0]
    >>> chart = graphic_frame.chart
    >>> chart_data = chart.chart_data
    >>> chart_data
    <pptx.chart.chart.ChartData instance at 0xdeadbeef1>
    >>> chart_data.update_from_xlsx_stream(xlsx_stream)


MS API Protocol
---------------

::

    >>> chart = prs.Slides(1).Shapes(1).Chart
    >>> chart_data = chart.ChartData
    >>> workbook = chart_data.Workbook
    >>> worksheet = Workbook.Worksheets(1)


ChartData objects
~~~~~~~~~~~~~~~~~

BreakLink
    Removes the link between the data for a chart and a Microsoft Excel
    workbook.

IsLinked
    True if the data for the chart is linked to an external Microsoft Excel
    workbook. Read-only Boolean.

Workbook
    Returns the workbook that contains the chart data associated with the
    chart. Read-only Object.


Code sketches
-------------

``ChartPart.xlsx_blob = blob``::

    @xlsx_blob.setter
    def xlsx_blob(self, blob):
        xlsx_part = self.xlsx_part
        if xlsx_part:
            xlsx_part.blob = blob
        else:
            xlsx_part = EmbeddedXlsxPart.new(blob, self.package)
            rId = self.relate_to(xlsx_part, RT.PACKAGE)
            externalData = self._element.get_or_add_externalData
            externalData.rId = rId

``@classmethod EmbeddedXlsxPart.new(cls, blob, package)``::

    partname = cls.next_partname(package)
    content_type = CT.SML_SHEET
    xlsx_part = EmbeddedXlsxPart(partname, content_type, blob, package)
    return xlsx_part


``ChartPart.add_or_replace_xlsx(xlsx_stream)``::

    xlsx_part = self.get_or_add_xlsx_part()
    xlsx_stream.seek(0)
    xlsx_bytes = xlsx_stream.read()
    xlsx_part.blob = xlsx_bytes


``ChartPart.xlsx_part``::

    externalData = self._element.externalData
    if externalData is None:
        raise ValueError("chart has no embedded worksheet")
    rId = externalData.rId
    xlsx_part = self.related_parts[rId]
    return xlsx_part

    # later ...

    xlsx_stream = BytesIO(xlsx_part.blob)
    xlsx_package = OpcPackage.open(xlsx_stream)
    workbook_part = xlsx_package.main_document


* Maybe can implement just a few Excel parts, enough to access and manipulate
  the data necessary. Like Workbook (start part I think) and Worksheet.

* What about linked rather than embedded Worksheet?


XML specimens
-------------

.. highlight:: xml

simple column chart::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
    <c:date1904 val="0"/>
    <c:lang val="en-US"/>
    <c:roundedCorners val="0"/>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
                 Requires="c14">
        <c14:style val="102"/>
      </mc:Choice>
      <mc:Fallback>
        <c:style val="2"/>
      </mc:Fallback>
    </mc:AlternateContent>
    <c:chart>
      <!-- 179 rows elided -->
    </c:chart>
    <c:txPr>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
        <a:pPr>
          <a:defRPr sz="1800"/>
        </a:pPr>
        <a:endParaRPr lang="en-US"/>
      </a:p>
    </c:txPr>
    <c:externalData r:id="rId1">
      <c:autoUpdate val="0"/>
    </c:externalData>
  </c:chartSpace>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_ChartSpace">
    <xsd:sequence>
      <xsd:element name="date1904"       type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="lang"           type="CT_TextLanguageID"    minOccurs="0"/>
      <xsd:element name="roundedCorners" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="style"          type="CT_Style"             minOccurs="0"/>
      <xsd:element name="clrMapOvr"      type="a:CT_ColorMapping"    minOccurs="0"/>
      <xsd:element name="pivotSource"    type="CT_PivotSource"       minOccurs="0"/>
      <xsd:element name="protection"     type="CT_Protection"        minOccurs="0"/>
      <xsd:element name="chart"          type="CT_Chart"/>
      <xsd:element name="spPr"           type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"           type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="externalData"   type="CT_ExternalData"      minOccurs="0"/>
      <xsd:element name="printSettings"  type="CT_PrintSettings"     minOccurs="0"/>
      <xsd:element name="userShapes"     type="CT_RelId"             minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ExternalData">
    <xsd:sequence>
      <xsd:element name="autoUpdate" type="CT_Boolean" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute ref="r:id" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Boolean">
    <xsd:attribute name="val" type="xsd:boolean" default="true"/>
  </xsd:complexType>
