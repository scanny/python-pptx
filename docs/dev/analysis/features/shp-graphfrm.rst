
Graphic Frame
=============

A *graphic frame* is a shape that contains a graphical object such as a table,
chart, or SmartArt object. While often referred to as a shape in common
parlance, a table or chart is not a shape. The graphic frame container is the
shape and the table, chart, or SmartArt object is a DrawingML (DML) object
within it.

A chart, for example, is a shared DML object. It can appear in a Word, Excel,
or PowerPoint file, virtually unchanged except for the container in which it
appears.

The graphical content is contained in the
``p:graphicFrame/a:graphic/a:graphicData`` grandchild element. The type of
graphical object contained is specified by an XML namespace contained in the
``uri`` attribute of the ``<a:graphicData>`` element. The graphical content may
appear directly in the ``<p:graphicFrame>`` element or may be in a separate
part related by an rId. XML for a table is embedded inline. Chart XML is stored
in a related Chart part.


Protocol
--------

::

    >>> shape = shapes.add_chart(style, type, x, y, cx, cy)
    >>> type(shape)
    <class 'pptx.shapes.graphfrm.GraphicFrame'>
    >>> shape.has_chart
    True
    >>> shape.has_table
    False
    >>> shape.chart
    <pptx.parts.chart.ChartPart object at 0x108c0e290>

    >>> shape = shapes.add_table(rows=2, cols=2, x, y, cx, cy)
    >>> type(shape)
    <class 'pptx.shapes.graphfrm.GraphicFrame'>
    >>> shape.has_chart
    False
    >>> shape.has_table
    True
    >>> shape.table
    <pptx.shapes.table.Table object at 0x108c0e310>


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
