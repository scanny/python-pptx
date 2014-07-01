
Chart Builder
=============

The purpose of the chart builder is to allow a developer to specify
a substantial number of the properties of a new chart and then add the chart
all in one go. All settings have a default, so a default chart can be inserted
very easily if that is desired. The developer only specifies items that differ
from the defaults.

This approach allows simpler generation code to be developed for a good part of
the chart XML and postpones the need for general-purpose modification grade
code for those elements.


Notes
-----

* [ ] will need an embedded package manager, should belong to package I think,
      like ImageParts.


Protocol
--------

::

    >>> from pptx.chart.builder import ChartBuilder
    >>> chart_builder = ChartBuilder(type=XL_CHART_TYPE.BAR_CLUSTERED)
    >>> chart_builder.categories = ('Foo', 'Bar', 'Baz')
    >>> chart_builder.add_series(1, 2, 3)

    >>> chart_graphfrm = chart_builder.add_chart(slide, x, y, cx, cy)
    ... OR ...
    >>> shape = shapes.add_chart_from_builder(chart_builder, x, y, cx, cy)


Protocol Notes
--------------

All parameters to ``ChartBuilder()`` are optional and are provided with
a default the same as the PowerPoint default.


Code sketch for ChartBuilder
----------------------------

::

    def add_chart(self, slide):
        embedded_package_part = self._create_worksheet() 
        chart_part = self.build_chart_part(embedded_package_part)
        graphic_frame = slide.shapes.add_graphic_frame_from_chart_part(
            chart_part, x, y, cx, cy
        )
        return graphic_frame

    def _create_worksheet(self):
        """
        Return an |EmbeddedPackage| object containing a spreadsheet populated
        with the data for this chart.
        """
        xlsx_file = StringIO()
        workbook = Workbook(xlsx_file, {'in_memory': True})
        self.populate_workbook(workbook)
        workbook.close()

    def _build_chart_part(self, embedded_chart_part):
        chart_part = ChartPart
        xlsx_rId = chart_part.relate_to(embedded_chart_part)
        chart_xml = self._generate_chart_xml(xlsx_rId)
        chart_part.blob = chart_xml
        return chart_part

    def _SlideShapeTree.add_graphic_frame_from_chart_part(self, chart_part):
        graphicFrame = self._add_chart_graphicFrame(x, y, cx, cy)
        graphic_frame = self._shape_factory(graphicFrame)
        return graphic_frame

    def _SlideShapeTree.add_chart_graphicFrame(self, x, y, cx, cy):
        shape_id, name = self._next_shape_credentials('Chart %d')
        graphicFrame = CT_GraphicalObjectFrame.new(
            CHART_URI, shape_id, name, x, y, cx, cy
        )
        self._spTree.append(graphicFrame)
        return graphicFrame


Enumerations
------------

XL_CHART_TYPE (XlChartType)
http://msdn.microsoft.com/en-us/library/office/ff838409.aspx


Specimen XML
------------

.. highlight:: xml

Table in a graphic frame::

    <p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="2" name="Table 1"/>
        <p:cNvGraphicFramePr>
          <a:graphicFrameLocks noGrp="1"/>
        </p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="1524000" y="1397000"/>
        <a:ext cx="6096000" cy="741680"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            ... remaining table XML ...
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>


Chart in a graphic frame::

    <p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="2" name="Chart 1"/>
        <p:cNvGraphicFramePr/>
        <p:nvPr>
          <p:extLst>
            <p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}">
              <p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
                         val="11227583"/>
            </p:ext>
          </p:extLst>
        </p:nvPr>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="1524000" y="1397000"/>
        <a:ext cx="6096000" cy="4064000"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                   r:id="rId2"/>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>


Related Schema Definitions
--------------------------

.. highlight:: xml

A ``<p:graphicFrame>`` element appears in a ``CT_GroupShape`` element,
typically a ``<p:spTree>`` (shape tree) element::

  <xsd:complexType name="CT_GroupShape">
    <xsd:sequence>
      <xsd:element name="nvGrpSpPr" type="CT_GroupShapeNonVisual"/>
      <xsd:element name="grpSpPr"   type="a:CT_GroupShapeProperties"/>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="sp"           type="CT_Shape"/>
        <xsd:element name="grpSp"        type="CT_GroupShape"/>
        <xsd:element name="graphicFrame" type="CT_GraphicalObjectFrame"/>
        <xsd:element name="cxnSp"        type="CT_Connector"/>
        <xsd:element name="pic"          type="CT_Picture"/>
        <xsd:element name="contentPart"  type="CT_Rel"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>


Graphic frame-related elements::

  <xsd:complexType name="CT_GraphicalObjectFrame">
    <xsd:sequence>
      <xsd:element name="nvGraphicFramePr" type="CT_GraphicalObjectFrameNonVisual"/>
      <xsd:element name="xfrm"             type="a:CT_Transform2D"/>
      <xsd:element ref="a:graphic"/>  <!-- type="CT_GraphicalObject" -->
      <xsd:element name="extLst"           type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="a:ST_BlackWhiteMode"/>
  </xsd:complexType>

  <xsd:complexType name="CT_GraphicalObjectFrameNonVisual">
    <xsd:sequence>
      <xsd:element name="cNvPr"             type="a:CT_NonVisualDrawingProps"/>
      <xsd:element name="cNvGraphicFramePr" type="a:CT_NonVisualGraphicFrameProperties"/>
      <xsd:element name="nvPr"              type="CT_ApplicationNonVisualDrawingProps"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GraphicalObject">
    <xsd:sequence>
      <xsd:element name="graphicData" type="CT_GraphicalObjectData"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GraphicalObjectData">
    <xsd:sequence>
      <xsd:any minOccurs="0" maxOccurs="unbounded" processContents="strict"/>
    </xsd:sequence>
    <xsd:attribute name="uri" type="xsd:token" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="hlinkClick" type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="hlinkHover" type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="id"     type="ST_DrawingElementId" use="required"/>
    <xsd:attribute name="name"   type="xsd:string"          use="required"/>
    <xsd:attribute name="descr"  type="xsd:string"          default=""/>
    <xsd:attribute name="hidden" type="xsd:boolean"         default="false"/>
    <xsd:attribute name="title"  type="xsd:string"          default=""/>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualGraphicFrameProperties">
    <xsd:sequence>
      <xsd:element name="graphicFrameLocks" type="CT_GraphicalObjectFrameLocking" minOccurs="0"/>
      <xsd:element name="extLst"            type="CT_OfficeArtExtensionList"      minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GraphicalObjectFrameLocking">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="noGrp"          type="xsd:boolean" default="false"/>
    <xsd:attribute name="noDrilldown"    type="xsd:boolean" default="false"/>
    <xsd:attribute name="noSelect"       type="xsd:boolean" default="false"/>
    <xsd:attribute name="noChangeAspect" type="xsd:boolean" default="false"/>
    <xsd:attribute name="noMove"         type="xsd:boolean" default="false"/>
    <xsd:attribute name="noResize"       type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ApplicationNonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="ph"          type="CT_Placeholder"      minOccurs="0"/>
      <xsd:group   ref="a:EG_Media"                              minOccurs="0"/>
      <xsd:element name="custDataLst" type="CT_CustomerDataList" minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="isPhoto"   type="xsd:boolean" default="false"/>
    <xsd:attribute name="userDrawn" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:group name="EG_Media">
    <xsd:choice>
      <xsd:element name="audioCd"       type="CT_AudioCD"/>
      <xsd:element name="wavAudioFile"  type="CT_EmbeddedWAVAudioFile"/>
      <xsd:element name="audioFile"     type="CT_AudioFile"/>
      <xsd:element name="videoFile"     type="CT_VideoFile"/>
      <xsd:element name="quickTimeFile" type="CT_QuickTimeFile"/>
    </xsd:choice>
  </xsd:group>
