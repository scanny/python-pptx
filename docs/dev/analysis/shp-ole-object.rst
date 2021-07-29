
Shape - OLE Object
==================

An embedded Excel worksheet can appear as a shape on a slide. The shape itself is a
graphics-frame, in which the worksheet "file" is embedded. The Excel binary appears as
an additional part in the package. Other file types such as DOCX and PDF can also be
embedded in a similar way.

Support for different file (object) types is OS dependent. For example, there is no
support for embedded PPTX objects on MacOS although there is for Windows. Excel (XLSX)
and Word (DOCX) are the two common denominators.


Proposed protocol
-----------------

::

    >>> from pptx.enum.shapes import PROG_ID
    >>> shapes = prs.slides[0].shapes
    >>> shape = shapes.add_ole_object(
            "worksheet.xlsx", PROG_ID.EXCEL, left, top, width, height
        )
    >>> shape
    <pptx.shapes.graphfrm.GraphicFrame instance at 0x00000008>
    >>> shape.shape_type
    MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
    >>> ole_format = shape.ole_format
    >>> ole_format
    <pptx.shapes.graphfrm._OleFormat instance at 0x00000010>
    >>> ole_format.blob
    <bytes instance at 0x00000018>
    >>> ole_format.prog_id
    'Excel.Sheet.12'
    >>> ole_format.show_as_icon
    True


PowerPoint UI behaviors
-----------------------

* **Insert a blank object.** PowerPoint allows a "new", blank OLE object to be inserted by
  selecting the type. The available types appear to be determined by a set of
  "registered" OLE server applications.

  This behavior seems quite bound to some user interaction to fill out the object and in
  any case is outside the abilities of `python-pptx` running on an arbitrary OS. This
  behavior will have no counterpart in `python-pptx`. Only the "insert-from-file" mode
  will be available in `python-pptx`.

* **Auto-detect `progId` from file.** When the "insert-from-file" option is chosen,
  PowerPoint automatically detects the `progId` (str OLE server identifier, intuitively
  "file-type") for the inserted file.

  This would be a lot of work for `python-pptx` to accomplish and in the end would be
  partial at best. The base case would be that `python-pptx` requires `progId` to be
  specified by the caller. `python-pptx` will provide an Enum that provides ready access
  to common str `progId` cases like Excel Worksheet, Word Document, and PowerPoint
  Presentation. Other cases could be resolved by experimentation and inspecting the XML
  to determine the appropriate str `progId` value.

* **Show as preview.** PowerPoint allows the option to show the embedded object as
  either a "preview" of the content or as an icon. The preview is an image (EMF for XLSX
  and perhaps other MS Office applications) and I expect it is provided by the OLE
  server application (Excel in the XLSX case) which renders to an image instead of
  rendering to the screen.

  `python-pptx` has no access to an OLE server and therefore cannot request this image.
  Only the "display as icon" mode will be supported in `python-pptx`. For this reason,
  the `python-pptx` call will have no `display_as_icon` parameter and that "value" will
  always be True.

* **User-selectable and partially-generated icon.** When "display-as-icon" is selected,
  PowerPoint generates a composite icon that combines a graphical icon and a caption.
  Together, these typically look like a file icon that appears on the user's desktop,
  except the file-name portion is a generic file-type name like "Microsoft Excel
  Workbook".

  The icon portion can be replaced with an arbitrary icon selected by the user (Windows
  only). The icon can be selected from a `.ico` file or an icon library or from the
  resource area of an `.exe` file.

  The resulting image in the PPTX package is a Windows Meta File (WMF/EMF), a vector
  format, perhaps to allow smooth scaling.

  In the general case (i.e. non-Windows OS), `python-pptx` has neither the option of
  extracting icons from an arbitrary icon or resource library or 

* **Link to file rather than embed.** When inserting an object from a file, a user can
  choose to link the selected file rather than embed it. This option appears in the
  Windows version only.

  The only difference in the slide XML is the appearance of a `<p:link>` element under
  `<p:oleObj>` rather than a `<p:embed>` element. In this case, the relationship
  referred to by `r:id` attribute of `p:oleObj` is an external link to the target file
  and no OLE object part is added to the PPTX package.

* **Auto-position and size.** In the PowerPoint UI, the OLE object is inserted at a
  fixed size and in a default location. Afterward the shape can be repositioned and
  resized to suit. In the API, the position and size are specified in the call.


MS API Protocol
---------------

::

    Shapes.AddOLEObject(
        Left, Top, Width, Height,
        ClassName,
        FileName,
        DisplayAsIcon,
        IconFileName,
        IconIndex,
        IconLabel,
        Link,
    )

* `Left, Top, Width` and `Height` specify the position and size of the graphic-frame
  object. They also specify the size of the icon image when `display_as_icon` is True
  and the size of the preview image when `display_as_icon` is False or omitted. These
  are all optional, and are determined by PowerPoint if not specified. I believe the
  shape is placed centered in the slide if the position is not specified and the size is
  determined by the icon or preview graphic if not specified.

  PowerPoint updates these values in the icon-image element to track those of the
  graphic-frame shape when the user changes the size of the graphic-frame shape, but it
  seems to ignore these values when they don't agree.

* `ClassName` identifies the "file-type" or more precisely, the OLE server (program)
  used to "open" the embedded or linked file. This is used only when a new, empty object
  is being added because otherwise PowerPoint derives this from the file specified with
  `FileName`

  This can be either the OLE long class name or the ProgID for the object that's to be
  created, but a class-name ends up being converted to a ProgID. Either the ClassName or
  FileName argument for the object must be specified, but not both. ClassName triggers
  the "insert-newly-created-object" mode and FileName triggers "insert-existing-object"
  mode.

* `DisplayAsIcon` (optional boolean) determines whether the OLE object will be displayed
  as an icon or as a "preview". The default is `False`.

* `IconFileName` allows the user to specify an *icon file* containing the icon to
  display when `DisplayAsIcon` is `True`. If not specified, a default icon for the OLE
  class is used. Note that this file can contain a collection of images, which is why
  the `IconIndex` parameter is available. These icon files are Windows specific and
  would not typically be found in other operating systems.

* `IconIndex` specifies	the index of the desired icon within `IconFileName`. The first
  icon in the file has the index number 0 (zero). If an icon with the given index number
  doesn't exist in IconFileName, the icon with the index number 1 (the second icon in
  the file) is used. The default value is 0 (zero).

* `IconLabel` is a str label (caption) to be displayed beneath the icon. By default,
  this is like "Microsoft Excel Worksheet". This caption is integrated into the
  specified "display-as-icon" image.

* `Link` is a boolean flag that determines whether the OLE object will be linked to the
  file from which it was created (rather than embedded). If you specified a value for
  ClassName, this argument must be msoFalse (linking is not an option in
  "insert-newly-created-object" mode).


Candiate protocol
-----------------

::

    SlideShapes.add_ole_object(
        object_file,
        prog_id,
        left,
        top,
        width=None,
        height=None,
        icon=None,
        link=False,
    )

`python-pptx` only supports adding an OLE object in "display-as-icon" mode. It has no
way of soliciting a preview image from an OLE server application, so that option is not
practical for us.

* `object_file` is the file containing the object to be inserted. It may be either a str
  path to the file or a (binary) file-like object (typically `io.Bytes`) containing the
  bytes of the file and implementing file-object semantics like `.read()` and `.seek()`.

* `prog_id` is a PROG_ID Enum member or str identifier like `"Excel.Sheet.12"`
  specifying the "type" of the object in terms of what application is used to "open" it.
  In Microsoft parlance, this identifies the OLE server called upon to operate on this
  object.

  The `pptx.enum.shapes.PROG_ID` enumeration defines these values for common cases like
  an Excel workbook, an Word document, or another PowerPoint presentation. Probably we
  should also include PDFs and any other common cases we can think of.

  A regular `str` value can be discovered by inspecting the XML of an example
  presentation and these will work just the same as a `PROG_ID` Enum value, allowing
  ready expansion to other OLE object types.

  I expect that a file of any type could be included, even if it doesn't have an OLE
  server application and it could then at least be accessed via `python-pptx`, although
  I don't suppose it would do anything useful from the PowerPoint UI. In any case, I
  don't believe it would raise an error and there wouldn't be anything we could (or
  would probably want) to do to stop someone from doing that.

* `left` and `top` are each an Emu object (or an int interpreted as Emu) and specify the
  position of the inserted-object shape.

* `width` and `height` are optional Emu/int objects and together specify the size of the
  graphic-frame object. Their use is not required and perhaps even discouraged unless
  the defaults of 1.00" (914400 EMU) wide and .84" (771480 EMU) tall do not suit for
  some reason, perhaps because the provided icon image is a non-standard size. The
  default size is that when a user inserts an object displayed as an icon in the
  PowerPoint UI and I at least have been unable to make it look better by resizing it.

* `icon` is an arbitrary image that appears in the graphic-frame object in lieu of the
  inserted object. It is optional, because a default icon is provided for each of the
  members of `PROG_ID` and this image need not be specified when `prog_id` is an
  instance of `PROG_ID`. Like an image object used in `SlideShapes.add_picture()`, this
  object can be either a `str` path or a file-like object (typically `io.BytesIO`)
  containing the image.

  This parameter is technically optional, but is required when `prog_id` is not an
  member of `PROG_ID` (because in that case we have no default icon available). The
  caller can always specify a custom icon image, even when inserting an object type
  available in `PROG_ID`. In that case, the image provided is used instead of the
  default icon.

* `link` is a boolean indicating the object should be linked rather than embedded.
  Linking probably only works in a Windows environment. This option may not be
  implemented in the initial release and this parameter will not appear in that case.

::

    >>> shapes = slide.shapes
    >>> embedded_xlsx_shape = shapes.AddOLEObject(left, top, width, height, file_name)


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

  <xsd:complexType name="CT_GraphicalObject">
    <xsd:sequence>
      <xsd:element name="graphicData" type="CT_GraphicalObjectData"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GraphicalObjectData">
    <xsd:sequence>
      <xsd:any minOccurs="0" maxOccurs="unbounded" processContents="strict"/>
    </xsd:sequence>
    <!-- contains "http://schemas.openxmlformats.org/presentationml/2006/ole" for an
         OLE-object graphic-frame -->
    <xsd:attribute name="uri" type="xsd:token" use="required"/>
  </xsd:complexType>

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
