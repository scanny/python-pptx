
SlideShapes.add_chart()
=======================

A chart is added to a slide similarly to adding any other shape. Note that
a chart is not itself a shape; the item returned by `.add_shape()` is
a |GraphicFrame| shape which contains a |Chart| object. The actual chart
object is accessed using the :attr:`chart` attribute on the graphic frame
that contains it.

Adding a chart requires three items, a chart type, the position and size
desired, and a |ChartData| object specifying the categories and series values
for the new chart.


Protocol
--------

Creating a new chart::

    >>> chart_data = ChartData()
    >>> chart_data.categories = 'Foo', 'Bar'
    >>> chart_data.add_series('Series 1', (1.2, 2.3))
    >>> chart_data.add_series('Series 2', (3.4, 4.5))

    >>> x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4)
    >>> graphic_frame = shapes.add_chart(
    >>>     XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    >>> )
    >>> chart = graphicFrame.chart


Specimen XML
------------

.. highlight:: xml

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
