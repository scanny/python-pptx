
Shape - Embedded Worksheet
==========================

An embedded Excel worksheet can appear as a shape on a slide. The shape itself is a
graphics-frame, in which the worksheet "file" is embedded. The Excel binary appears as
an additional part in the package. Other file types such as PPTX can also be embedded in
a similar way.


Spike approach
--------------

* [ ] form subtree literal starting with hardcode and advancing to using interpolation.
* [ ] insert shape subtree from literal
* [ ] write relationship

Refine for minimality

* [ ] Try dropping creationId extension
* [ ] Try dropping tracking extension
* [ ] Try dropping the fallback, but maybe that's okay

  Looks like there is a link to the .emf file for each embedding too.

* [ ] driver with one blank slide, builder column chart single series.
* [ ] Can edit Excel data in resulting file
* [ ] resulting file has an embedded Excel package of expected name
* [ ] XML uses worksheet references, not just cached data


Proposed protocol
-----------------

::

    >>> shapes = prs.slides[0].shapes
    >>> embedded_xlsx_shape = shapes.add_embedded_xlsx_shape(
            left, top, width, height, file_name
        )
    >>> embedded_xlsx_shape
    <pptx.shape.GraphicFrame instance at 0xdeadbeef1>
    >>> embedded_xlsx_shape.shape_type
    MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
    >>> ole_format = embedded_xlsx_shape.ole_format
    >>> ole_format
    <...OleFormat instance at 0x00000008>
    >>> old_format.prog_id
    'Excel.Sheet.12'
    >>> old_format.blob
    <bytes instance at 0x00000010>


MS API Protocol
---------------

::

    >>> shapes = slide.shapes
    >>> embedded_xlsx_shape = shapes.AddOLEObject(left, top, width, height, file_name)


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


relationships::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="x" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image4.emf"/>
    <Relationship Id="x" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
    <Relationship Id="x" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/>
    <Relationship Id="x" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout14.xml"/>
    <Relationship Id="x" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
  </Relationships>

simple column chart::

  <p:graphicFrame>
    <p:nvGraphicFramePr>
      <p:cNvPr id="2" name="Object 1">
        <a:extLst>
          <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
            <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{9DA7C2C3-4766-419F-9ED0-2856E43424DD}"/>
          </a:ext>
        </a:extLst>
      </p:cNvPr>
      <p:cNvGraphicFramePr>
        <a:graphicFrameLocks noChangeAspect="1"/>
      </p:cNvGraphicFramePr>
      <p:nvPr>
        <p:extLst>
          <p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}">
            <p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="2099550745"/>
          </p:ext>
        </p:extLst>
      </p:nvPr>
    </p:nvGraphicFramePr>
    <p:xfrm>
      <a:off x="1792101" y="2202989"/>
      <a:ext cx="659686" cy="1371600"/>
    </p:xfrm>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
        <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v">
            <p:oleObj spid="_x0000_s1058" name="Worksheet" showAsIcon="1" r:id="rId4" imgW="381148" imgH="792690" progId="Excel.Sheet.12">
              <p:embed/>
            </p:oleObj>
          </mc:Choice>
          <mc:Fallback>
            <p:oleObj name="Worksheet" showAsIcon="1" r:id="rId4" imgW="381148" imgH="792690" progId="Excel.Sheet.12">
              <p:embed/>
              <p:pic>
                <p:nvPicPr>
                  <p:cNvPr id="0" name=""/>
                  <p:cNvPicPr/>
                  <p:nvPr/>
                </p:nvPicPr>
                <p:blipFill>
                  <a:blip r:embed="rId5"/>
                  <a:stretch>
                    <a:fillRect/>
                  </a:stretch>
                </p:blipFill>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="1792101" y="2202989"/>
                    <a:ext cx="659686" cy="1371600"/>
                  </a:xfrm>
                  <a:prstGeom prst="rect">
                    <a:avLst/>
                  </a:prstGeom>
                </p:spPr>
              </p:pic>
            </p:oleObj>
          </mc:Fallback>
        </mc:AlternateContent>
      </a:graphicData>
    </a:graphic>
  </p:graphicFrame>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:element name="oleObj" type="CT_OleObject"/>

  <xsd:complexType name="CT_OleObject">
    <xsd:sequence>
      <xsd:choice minOccurs="1" maxOccurs="1">
        <xsd:element name="embed" type="CT_OleObjectEmbed"/>
        <xsd:element name="link" type="CT_OleObjectLink"/>
      </xsd:choice>
      <xsd:element name="pic" type="CT_Picture" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="spid" type="a:ST_ShapeID" use="optional"/>
    <xsd:attribute name="name" type="xsd:string" use="optional" default=""/>
    <xsd:attribute name="showAsIcon" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute ref="r:id" use="optional"/>
    <xsd:attribute name="imgW" type="a:ST_PositiveCoordinate32" use="optional"/>
    <xsd:attribute name="imgH" type="a:ST_PositiveCoordinate32" use="optional"/>
    <xsd:attribute name="progId" type="xsd:string" use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_OleObjectEmbed">
    <xsd:sequence>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="followColorScheme" type="ST_OleObjectFollowColorScheme" use="optional"
      default="none"/>
  </xsd:complexType>
