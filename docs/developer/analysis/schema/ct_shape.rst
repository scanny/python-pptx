============
``CT_Shape``
============

.. highlight:: xml

.. csv-table::
   :header-rows: 0
   :stub-columns: 1
   :widths: 15, 50

   Schema Name  , CT_Shape
   Spec Name    , Shape
   Tag(s)       , p:sp
   Namespace    , presentationml (pml.xsd)
   Schema Line  , 1209
   Spec Section , 19.3.1.43


Example
=======

::

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


Minimal text box ``sp`` shape
=============================

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="9" name="Text Box 8"/>
      <p:cNvSpPr txBox="1"/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="9999" y="9999"/>
        <a:ext cx="999999" cy="999999"/>
      </a:xfrm>
    </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p>
            <a:r>
              <a:t>This is text in a text box</a:t>
            </a:r>
          </a:p>
        </p:txBody>
  </p:sp>


Schema-minimal ``sp`` shape
===========================

The following XML represents the minimum valid ``<p:sp>`` element required by
the schema. Note that in general schema-minimal elements are not guaranteed to
be semantically valid and attempting to load a presentation that contains one
will often trigger a load error by the PowerPoint® client. The schema-minimal
element is useful, however, for understanding the duties of a constructor of
that element::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="9" name="Text Box 8"/>
      <p:cNvSpPr/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr/>
  </p:sp>


Analysis
========

The ``CT_Shape`` (``<p:sp>``) element is multi-purpose. One of the shape types
it is used for is a text box.


attributes
^^^^^^^^^^

================  ===  ==================  ========
name              use  type                default
================  ===  ==================  ========
useBgFill          ?   xsd:boolean         false
================  ===  ==================  ========


child elements
^^^^^^^^^^^^^^

======  ===  ======================  ========
name     #   type                    line
======  ===  ======================  ========
nvSpPr   1   CT_ShapeNonVisual       1201
spPr     1   CT_ShapeProperties      2210 dml
style    ?   CT_ShapeStyle           2245 dml
txBody   ?   |CT_TextBody|           2640 dml
extLst   ?   CT_ExtensionListModify  775>767
======  ===  ======================  ========

.. |CT_TextBody| replace:: :doc:`ct_textbody`


Spec text
^^^^^^^^^

   This element specifies the existence of a single shape. A shape can either
   be a preset or a custom geometry, defined using the DrawingML framework. In
   addition to a geometry each shape can have both visual and non-visual
   properties attached. Text and corresponding styling information can also be
   attached to a shape. This shape is specified along with all other shapes
   within either the shape tree or group shape elements.


Schema excerpt
^^^^^^^^^^^^^^

::

  <xsd:complexType name="CT_Shape">
    <xsd:sequence>
      <xsd:element name="nvSpPr" type="CT_ShapeNonVisual"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties"   minOccurs="1" maxOccurs="1"/>
      <xsd:element name="style"  type="a:CT_ShapeStyle"        minOccurs="0" maxOccurs="1"/>
      <xsd:element name="txBody" type="a:CT_TextBody"          minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="useBgFill" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeNonVisual">
    <xsd:sequence>
      <xsd:element name="cNvPr"   type="a:CT_NonVisualDrawingProps"       minOccurs="1" maxOccurs="1"/>
      <xsd:element name="cNvSpPr" type="a:CT_NonVisualDrawingShapeProps"  minOccurs="1" maxOccurs="1"/>
      <xsd:element name="nvPr" type="CT_ApplicationNonVisualDrawingProps" minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"    type="CT_Transform2D"            minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_Geometry"                               minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_FillProperties"                         minOccurs="0" maxOccurs="1"/>
      <xsd:element name="ln"      type="CT_LineProperties"         minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_EffectProperties"                       minOccurs="0" maxOccurs="1"/>
      <xsd:element name="scene3d" type="CT_Scene3D"                minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sp3d"    type="CT_Shape3D"                minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeStyle">
    <xsd:sequence>
      <xsd:element name="lnRef"     type="CT_StyleMatrixReference" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="fillRef"   type="CT_StyleMatrixReference" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="effectRef" type="CT_StyleMatrixReference" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="fontRef"   type="CT_FontReference"        minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_TextBody">
    <xsd:sequence>
      <xsd:element name="bodyPr"   type="CT_TextBodyProperties" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="lstStyle" type="CT_TextListStyle"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="p"        type="CT_TextParagraph"      minOccurs="1" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ExtensionListModify">
    <xsd:sequence>
      <xsd:group ref="EG_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="mod" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <!-- Supporting elements -->
  
  <xsd:complexType name="CT_NonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="hlinkClick" type="CT_Hyperlink"              minOccurs="0" maxOccurs="1"/>
      <xsd:element name="hlinkHover" type="CT_Hyperlink"              minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"     type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="id"     type="ST_DrawingElementId" use="required"/>
    <xsd:attribute name="name"   type="xsd:string"          use="required"/>
    <xsd:attribute name="descr"  type="xsd:string"          use="optional" default=""/>
    <xsd:attribute name="hidden" type="xsd:boolean"         use="optional" default="false"/>
    <xsd:attribute name="title"  type="xsd:string"          use="optional" default=""/>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualDrawingShapeProps">
    <xsd:sequence>
      <xsd:element name="spLocks" type="CT_ShapeLocking"           minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="txBox" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ApplicationNonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="ph"          type="CT_Placeholder"      minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="a:EG_Media"                              minOccurs="0" maxOccurs="1"/>
      <xsd:element name="custDataLst" type="CT_CustomerDataList" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"    minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="isPhoto"   type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="userDrawn" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Transform2D">
    <xsd:sequence>
      <xsd:element name="off" type="CT_Point2D" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="ext" type="CT_PositiveSize2D" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="rot" type="ST_Angle" use="optional" default="0"/>
    <xsd:attribute name="flipH" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="flipV" type="xsd:boolean" use="optional" default="false"/>
  </xsd:complexType>

  <xsd:group name="EG_Geometry">
    <xsd:choice>
      <xsd:element name="custGeom" type="CT_CustomGeometry2D" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="prstGeom" type="CT_PresetGeometry2D" minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_CustomGeometry2D">
    <xsd:sequence>
      <xsd:element name="avLst"   type="CT_GeomGuideList"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="gdLst"   type="CT_GeomGuideList"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="ahLst"   type="CT_AdjustHandleList"   minOccurs="0" maxOccurs="1"/>
      <xsd:element name="cxnLst"  type="CT_ConnectionSiteList" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="rect"    type="CT_GeomRect"           minOccurs="0" maxOccurs="1"/>
      <xsd:element name="pathLst" type="CT_Path2DList"         minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_PresetGeometry2D">
    <xsd:sequence>
      <xsd:element name="avLst" type="CT_GeomGuideList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="prst" type="ST_ShapeType" use="required"/>
  </xsd:complexType>

  <xsd:group name="EG_FillProperties">
    <xsd:choice>
      <xsd:element name="noFill"    type="CT_NoFillProperties"         minOccurs="1" maxOccurs="1"/>
      <xsd:element name="solidFill" type="CT_SolidColorFillProperties" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="gradFill"  type="CT_GradientFillProperties"   minOccurs="1" maxOccurs="1"/>
      <xsd:element name="blipFill"  type="CT_BlipFillProperties"       minOccurs="1" maxOccurs="1"/>
      <xsd:element name="pattFill"  type="CT_PatternFillProperties"    minOccurs="1" maxOccurs="1"/>
      <xsd:element name="grpFill"   type="CT_GroupFillProperties"      minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

  <xsd:complexType name="CT_LineProperties">
    <xsd:sequence>
      <xsd:group   ref="EG_LineFillProperties"                     minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_LineDashProperties"                     minOccurs="0" maxOccurs="1"/>
      <xsd:group   ref="EG_LineJoinProperties"                     minOccurs="0" maxOccurs="1"/>
      <xsd:element name="headEnd" type="CT_LineEndProperties"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="tailEnd" type="CT_LineEndProperties"      minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="w"    type="ST_LineWidth"    use="optional"/>
    <xsd:attribute name="cap"  type="ST_LineCap"      use="optional"/>
    <xsd:attribute name="cmpd" type="ST_CompoundLine" use="optional"/>
    <xsd:attribute name="algn" type="ST_PenAlignment" use="optional"/>
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
      <xsd:element name="effectLst" type="CT_EffectList"      minOccurs="1" maxOccurs="1"/>
      <xsd:element name="effectDag" type="CT_EffectContainer" minOccurs="1" maxOccurs="1"/>
    </xsd:choice>
  </xsd:group>

