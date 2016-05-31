
Shapes - In General
===================

Each visual element in a PowerPoint presentation is a *shape*. A shape
appears on the "canvas" of a slide, which includes the various types of
*master*. Within a slide, shapes appear in a *shape tree*, corresponding to
an ``<p:spTree>`` element.

The following table summarizes the six shape types:

============  ====================
shape type    element
============  ====================
auto shape    ``<p:sp>``
group shape   ``<p:grpSp>``
graphicFrame  ``<p:graphicFrame>``
connector     ``<p:cxnSp>``
picture       ``<p:pic>``
content part  ``<p:contentPart>``
============  ====================

Some of these shape types have important sub-types. For example,
a placeholder, a text box, and a preset geometry shape such as a circle, are
all defined with an ``<p:sp>`` element.


``<p:sp>`` shape elements
-------------------------

The ``<p:sp>`` element is used for three types of shape: placeholder, text
box, and geometric shapes. A geometric shape with preset geometry is referred
to as an *auto shape*. Placeholder shapes are documented on the
:ref:`placeholder` page. Auto shapes are documented on the :ref:`autoshape`
page.

Geometric shapes are the familiar shapes that may be placed on a slide such
as a rectangle or an ellipse. In the PowerPoint UI they are simply called
shapes. There are two types of geometric shapes, preset geometry shapes and
custom geometry shapes.


``Shape.id`` and ``Shape.name``
-------------------------------

``Shape.id`` is read-only and is assigned by python-pptx when necessary.

Proposed protocol::

  >>> shape.id
  42
  >>> shape.name
  u'Picture 5'
  >>> shape.name = 'T501 - Foo; B. Baz; 2014'
  >>> shape.name
  u'T501 - Foo; B. Baz; 2014'


:attr:`Shape.rotation`
----------------------

Read/write float degrees of clockwise rotation. Negative values can be used
for counter-clockwise rotation.

*XML Semantics*
    ST_Angle is an integer value, 60,000 to each degree. PowerPoint appears
    to only ever use positive values. Oddly, the UI uses positive values for
    counter-clockwise rotation while the XML uses positive increase for
    clockwise rotation.

*PowerPoint behavior*
    It appears graphic frame shapes can't be rotated. AutoShape, group,
    connector, and picture all rotate fine.


Proposed protocol::

  >>> shape.rotation
  0.0
  >>> shape.rotation = 45.2
  >>> shape.rotation
  45.2


Math::

    def rot_from_angle(value):
        """
        Return positive integer rotation in 60,000ths of a degree
        corresponding to *value* expressed as a float number of degrees.
        """
        DEGREE_INCREMENTS = 60000
        THREE_SIXTY = 360 * DEGREE_INCREMENTS
        # modulo normalizes negative and >360 degree values
        return int(round(value * DEGREE_INCREMENTS)) % THREE_SIXTY


Specimen XML
------------

.. highlight:: xml

Geometric shape (rounded rectangle)::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="3" name="Rounded Rectangle 2"/>
      <p:cNvSpPr/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="760096" y="562720"/>
        <a:ext cx="2520824" cy="914400"/>
      </a:xfrm>
      <a:prstGeom prst="roundRect">
        <a:avLst>
          <a:gd name="adj" fmla="val 30346"/>
        </a:avLst>
      </a:prstGeom>
    </p:spPr>
    <p:style>
      <a:lnRef idx="1">
        <a:schemeClr val="accent1"/>
      </a:lnRef>
      <a:fillRef idx="3">
        <a:schemeClr val="accent1"/>
      </a:fillRef>
      <a:effectRef idx="2">
        <a:schemeClr val="accent1"/>
      </a:effectRef>
      <a:fontRef idx="minor">
        <a:schemeClr val="lt1"/>
      </a:fontRef>
    </p:style>
    <p:txBody>
      <a:bodyPr rtlCol="0" anchor="ctr"/>
      <a:lstStyle/>
      <a:p>
        <a:pPr algn="ctr"/>
        <a:r>
          <a:rPr lang="en-US" dirty="0" smtClean="0"/>
          <a:t>This is text inside a rounded rectangle</a:t>
        </a:r>
        <a:endParaRPr lang="en-US" dirty="0"/>
      </a:p>
    </p:txBody>
  </p:sp>


Default textbox shape as inserted by PowerPoint::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="2" name="TextBox 1"/>
      <p:cNvSpPr txBox="1"/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="1997289" y="2529664"/>
        <a:ext cx="2390398" cy="369332"/>
      </a:xfrm>
      <a:prstGeom prst="rect">
        <a:avLst/>
      </a:prstGeom>
      <a:noFill/>
    </p:spPr>
    <p:txBody>
      <a:bodyPr wrap="none" rtlCol="0">
        <a:spAutoFit/>
      </a:bodyPr>
      <a:lstStyle/>
      <a:p>
        <a:r>
          <a:rPr lang="en-US" dirty="0" smtClean="0"/>
          <a:t>This is text in a text box</a:t>
        </a:r>
        <a:endParaRPr lang="en-US" dirty="0"/>
      </a:p>
    </p:txBody>
  </p:sp>


Group shape (some contents elided for size)::

  <p:grpSp>
    <p:nvGrpSpPr>
      <p:cNvPr id="4" name="Group 3"/>
      <p:cNvGrpSpPr/>
      <p:nvPr/>
    </p:nvGrpSpPr>
    <p:grpSpPr>
      <a:xfrm>
        <a:off x="2438400" y="2971800"/>
        <a:ext cx="4267200" cy="914400"/>
        <a:chOff x="2438400" y="2971800"/>
        <a:chExt cx="4267200" cy="914400"/>
      </a:xfrm>
    </p:grpSpPr>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="2" name="Rectangle 1"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <!-- some contents elided -->
    </p:sp>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="3" name="Oval 2"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <!-- some contents elided -->
    </p:sp>
  </p:grpSp>


Graphical object (e.g. table, chart) in a graphic frame::

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
        <!-- graphical object XML or ref goes here -->
      </a:graphicData>
    </a:graphic>
  </p:graphicFrame>


Connector shape::

  <p:cxnSp>
    <p:nvCxnSpPr>
      <p:cNvPr id="6" name="Straight Connector 5"/>
      <p:cNvCxnSpPr/>
      <p:nvPr/>
    </p:nvCxnSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="3131840" y="3068960"/>
        <a:ext cx="2736304" cy="0"/>
      </a:xfrm>
      <a:prstGeom prst="line">
        <a:avLst/>
      </a:prstGeom>
    </p:spPr>
    <p:style>
      <a:lnRef idx="2">
        <a:schemeClr val="accent1"/>
      </a:lnRef>
      <a:fillRef idx="0">
        <a:schemeClr val="accent1"/>
      </a:fillRef>
      <a:effectRef idx="1">
        <a:schemeClr val="accent1"/>
      </a:effectRef>
      <a:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </a:fontRef>
    </p:style>
  </p:cxnSp>


Picture shape::

  <p:pic>
    <p:nvPicPr>
      <p:cNvPr id="6" name="Picture 5" descr="python-logo.gif"/>
      <p:cNvPicPr>
        <a:picLocks noChangeAspect="1"/>
      </p:cNvPicPr>
      <p:nvPr/>
    </p:nvPicPr>
    <p:blipFill>
      <a:blip r:embed="rId2"/>
      <a:stretch>
        <a:fillRect/>
      </a:stretch>
    </p:blipFill>
    <p:spPr>
      <a:xfrm>
        <a:off x="5580112" y="1988840"/>
        <a:ext cx="2679700" cy="901700"/>
      </a:xfrm>
      <a:prstGeom prst="rect">
        <a:avLst/>
      </a:prstGeom>
      <a:ln>
        <a:solidFill>
          <a:schemeClr val="bg1">
            <a:lumMod val="85000"/>
          </a:schemeClr>
        </a:solidFill>
      </a:ln>
    </p:spPr>
  </p:pic>


Resources
---------

* `DrawingML Shapes`_ on officeopenxml.com

.. _DrawingML Shapes:
   http://officeopenxml.com/drwShape.php

* `Shape Object MSDN page`_

.. _Shape Object MSDN page:
   http://msdn.microsoft.com/en-us/library/office/ff744177(v=office.14).aspx

* `MsoShapeType Enumeration`_

.. _MsoShapeType Enumeration:
   http://msdn.microsoft.com/en-us/library/office/aa432678(v=office.14).aspx


Schema excerpt
--------------

::

  <xsd:complexType name="CT_Shape">
    <xsd:sequence>
      <xsd:element name="nvSpPr" type="CT_ShapeNonVisual"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties"/>
      <xsd:element name="style"  type="a:CT_ShapeStyle"        minOccurs="0"/>
      <xsd:element name="txBody" type="a:CT_TextBody"          minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="useBgFill" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeNonVisual">
    <xsd:sequence>
      <xsd:element name="cNvPr"   type="a:CT_NonVisualDrawingProps"/>
      <xsd:element name="cNvSpPr" type="a:CT_NonVisualDrawingShapeProps"/>
      <xsd:element name="nvPr"    type="CT_ApplicationNonVisualDrawingProps"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"                type="CT_Transform2D"            minOccurs="0"/>
      <xsd:group   ref ="EG_Geometry"                                          minOccurs="0"/>
      <xsd:group   ref ="EG_FillProperties"                                    minOccurs="0"/>
      <xsd:element name="ln"                  type="CT_LineProperties"         minOccurs="0"/>
      <xsd:group   ref ="EG_EffectProperties"                                  minOccurs="0"/>
      <xsd:element name="scene3d"             type="CT_Scene3D"                minOccurs="0"/>
      <xsd:element name="sp3d"                type="CT_Shape3D"                minOccurs="0"/>
      <xsd:element name="extLst"              type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeStyle">
    <xsd:sequence>
      <xsd:element name="lnRef"     type="CT_StyleMatrixReference"/>
      <xsd:element name="fillRef"   type="CT_StyleMatrixReference"/>
      <xsd:element name="effectRef" type="CT_StyleMatrixReference"/>
      <xsd:element name="fontRef"   type="CT_FontReference"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextBody">
    <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle" minOccurs="0"/>
      <xsd:element name="p"        type="CT_TextParagraph" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ExtensionListModify">
    <xsd:sequence>
      <xsd:group ref="EG_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="mod" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <!-- Supporting elements -->
  
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

  <xsd:complexType name="CT_NonVisualDrawingShapeProps">
    <xsd:sequence>
      <xsd:element name="spLocks" type="CT_ShapeLocking"           minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="txBox" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ApplicationNonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="ph"          type="CT_Placeholder"      minOccurs="0"/>
      <xsd:group   ref ="a:EG_Media"                             minOccurs="0"/>
      <xsd:element name="custDataLst" type="CT_CustomerDataList" minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="isPhoto"   type="xsd:boolean" default="false"/>
    <xsd:attribute name="userDrawn" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Transform2D">
    <xsd:sequence>
      <xsd:element name="off" type="CT_Point2D"        minOccurs="0"/>
      <xsd:element name="ext" type="CT_PositiveSize2D" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="rot"   type="ST_Angle"    default="0"/>
    <xsd:attribute name="flipH" type="xsd:boolean" default="false"/>
    <xsd:attribute name="flipV" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:group name="EG_Geometry">
    <xsd:choice>
      <xsd:element name="custGeom" type="CT_CustomGeometry2D"/>
      <xsd:element name="prstGeom" type="CT_PresetGeometry2D"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_CustomGeometry2D">
    <xsd:sequence>
      <xsd:element name="avLst"   type="CT_GeomGuideList"      minOccurs="0"/>
      <xsd:element name="gdLst"   type="CT_GeomGuideList"      minOccurs="0"/>
      <xsd:element name="ahLst"   type="CT_AdjustHandleList"   minOccurs="0"/>
      <xsd:element name="cxnLst"  type="CT_ConnectionSiteList" minOccurs="0"/>
      <xsd:element name="rect"    type="CT_GeomRect"           minOccurs="0"/>
      <xsd:element name="pathLst" type="CT_Path2DList"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_PresetGeometry2D">
    <xsd:sequence>
      <xsd:element name="avLst" type="CT_GeomGuideList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="prst" type="ST_ShapeType" use="required"/>
  </xsd:complexType>

  <xsd:group name="EG_FillProperties">
    <xsd:choice>
      <xsd:element name="noFill"    type="CT_NoFillProperties"/>
      <xsd:element name="solidFill" type="CT_SolidColorFillProperties"/>
      <xsd:element name="gradFill"  type="CT_GradientFillProperties"/>
      <xsd:element name="blipFill"  type="CT_BlipFillProperties"/>
      <xsd:element name="pattFill"  type="CT_PatternFillProperties"/>
      <xsd:element name="grpFill"   type="CT_GroupFillProperties"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_LineProperties">
    <xsd:sequence>
      <xsd:group   ref="EG_LineFillProperties"                     minOccurs="0"/>
      <xsd:group   ref="EG_LineDashProperties"                     minOccurs="0"/>
      <xsd:group   ref="EG_LineJoinProperties"                     minOccurs="0"/>
      <xsd:element name="headEnd" type="CT_LineEndProperties"      minOccurs="0"/>
      <xsd:element name="tailEnd" type="CT_LineEndProperties"      minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="w"    type="ST_LineWidth"/>
    <xsd:attribute name="cap"  type="ST_LineCap"/>
    <xsd:attribute name="cmpd" type="ST_CompoundLine"/>
    <xsd:attribute name="algn" type="ST_PenAlignment"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Point2D">
    <xsd:attribute name="x" type="ST_Coordinate" use="required"/>
    <xsd:attribute name="y" type="ST_Coordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PositiveSize2D">
    <xsd:attribute name="cx" type="ST_PositiveCoordinate" use="required"/>
    <xsd:attribute name="cy" type="ST_PositiveCoordinate" use="required"/>
  </xsd:complexType>

  <xsd:group name="EG_EffectProperties">
    <xsd:choice>
      <xsd:element name="effectLst" type="CT_EffectList"/>
      <xsd:element name="effectDag" type="CT_EffectContainer"/>
    </xsd:choice>
  </xsd:group>

  <xsd:group name="EG_Media">
    <xsd:choice>
      <xsd:element name="audioCd"       type="CT_AudioCD"/>
      <xsd:element name="wavAudioFile"  type="CT_EmbeddedWAVAudioFile"/>
      <xsd:element name="audioFile"     type="CT_AudioFile"/>
      <xsd:element name="videoFile"     type="CT_VideoFile"/>
      <xsd:element name="quickTimeFile" type="CT_QuickTimeFile"/>
    </xsd:choice>
  </xsd:group>

  <xsd:simpleType name="ST_DrawingElementId">
    <xsd:restriction base="xsd:unsignedInt"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Angle">
    <xsd:restriction base="xsd:int"/>
  </xsd:simpleType>
