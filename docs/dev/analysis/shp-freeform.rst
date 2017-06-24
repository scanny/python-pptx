
Freeform Shapes
===============

PowerPoint AutoShapes, such as rectangles, circles, and triangles, have
a *preset* geometry. The user simply picks from a selection of some 160
pre-defined shapes to place one on the slide, at which point they can be
scaled. Many shapes allow their geometry to be adjusted in certain ways using
adjustment points (by dragging the small yellow diamonds).

PowerPoint also supports a much less commonly used but highly flexible sort
of shape having *custom geometry*. Whereas a shape with preset geometry has
that geometry defined internally to PowerPoint and simply referred to by
a keyname (such as 'rect'), a shape with custom geometry spells out the
custom geometry step-by-step in an XML subtree, using a sequence of moveTo,
lineTo, and other drawing commands.


Scope
-----

* Move to, line to, and close, initially. Also something about multiple
  paths.
* Line style (solid, dashed, etc.)
* Fill pattern (solid, dashed, etc.)
* Shape shadow
* What about stroke on path? Seems like this could just be determined with
  None on like fill.


TODO
----

* Get a test opc extract going for iterative manual editing to see how things
  behave. Start with a single line segment and work out from there.

  + [ ] See what happens if there's no `a:moveTo` element at the start of
        a path.

  + [ ] Derive all the rules for "subtraction". It's clear that a second
        drawing sequence in a path that is completely contained inside
        a prior drawing sequence appears as a "cutout". What's not clear is
        what happens when the second sequence "overlaps" the prior sequence,
        having some vertices inside and some outside.

  + [ ] Confirm: This subtraction behavior does not occur when the two
        drawing sequences are in separate paths, even within the same shape.

  + [ ] What happens if you don't close path 1 and then start path 2 with
        an `a:moveTo` element?

  + [ ] What happens if you do a shape with an `a:moveTo` in the middle
        (producing a gap in the outline) but then close the shape? Does it
        still get a fill or is it considered open then?

  + [ ] What's up with z-order in paths? Do all lines show through one
        another or is there some sort of stacking behavior?

* Work out scaling strategy. Offset (position) is determined by `a:xfrm`, so
  working with a local coordinate space makes sense, but that would be top,
  left == (0, 0). Has to be integers I'll bet. I expect they need to fit in
  a long int, so maybe limited to 4E9.

* Experiment: see if PowerPoint likes it okay if we leave out all the
  "guides" and connection points, i.e. if the `a:gdLst` and `a:cxnLst`
  elements are empty.


XML Semantics
-------------

* The `a:pathLst` element contains zero or more `a:path` elements, each of
  which specify a *path*.

* A path is composed of a sequence of the following possible elements:

  + `a:moveTo`
  + `a:lnTo`
  + `a:arcTo`
  + `a:quadBezTo`
  + `a:cubicBezTo`
  + `a:close`

  A path may begin with an `a:moveTo` element. This essentially locates the
  starting location of the "pen". Each subsequent drawing command extends the
  shape by adding a line segment. If the path does not begin with an
  `a:moveTo` element, the origin (0, 0) is used as the initial pen location.

  A path can be open or closed. If an `a:close` element is added, a straight
  line segment is drawn from the current pen location to the initial location
  of the drawing sequence and the shape appears with a fill. If the pen is
  already at the starting location, no additional line segment appears. If no
  `a:close` element is added, the shape remains "open" and only the path
  appears (no fill).

  A path can contain more than one drawing sequence, i.e. one sequence can be
  "closed" and another sequence started. If a subsequent drawing sequence is
  entirely enclosed within a prior sequence, it appears as a "cutout", or an
  interior boundary. This behavior does not occur when the two drawing
  sequences are in separate paths, even within the same shape.

  The pen can be "lifted" using an `a:moveTo` element, in which case no line
  segment is drawn between the prior location and the new location. This can
  be used to produce a discontinuous outline.

  A path has boolean a `stroke` attribute (default True) that specifies
  whether a line should appear on the path.

* The `a:pathLst` element can contain multiple `a:path` elements. In this
  case, each path is essentially a "sub-shape", such as a shape that depicts
  the islands of Hawaii. 

  If a prior path is not closed, its end point path will be connected to the
  first point of the subsequent path.

  The paths within a shape all have the same z-position, i.e. they appear on
  a single plane such that all outlines appear, even when they intersect.
  There is no cropping behavior such as occurs for individual shapes on
  a slide.

  <xsd:complexType name="CT_Path2D">
    <xsd:choice minOccurs="0" maxOccurs="unbounded">
      <xsd:element name="close"      type="CT_Path2DClose"/>
      <xsd:element name="moveTo"     type="CT_Path2DMoveTo"/>
      <xsd:element name="lnTo"       type="CT_Path2DLineTo"/>
      <xsd:element name="arcTo"      type="CT_Path2DArcTo"/>
      <xsd:element name="quadBezTo"  type="CT_Path2DQuadBezierTo"/>
      <xsd:element name="cubicBezTo" type="CT_Path2DCubicBezierTo"/>
    </xsd:choice>
    <xsd:attribute name="w"           type="ST_PositiveCoordinate" default="0"/>
    <xsd:attribute name="h"           type="ST_PositiveCoordinate" default="0"/>
    <xsd:attribute name="fill"        type="ST_PathFillMode"       default="norm"/>
    <xsd:attribute name="stroke"      type="xsd:boolean"           default="true"/>
    <xsd:attribute name="extrusionOk" type="xsd:boolean"           default="true"/>
  </xsd:complexType>


Coordinate system
~~~~~~~~~~~~~~~~~

* Each path has its own local coordinate system, distinct both from the
  *shape* coordinate system and the coordinate systems of the other paths in
  the shape.

* The x and y extents of a path coordinate system are specified by the `w`
  and `h` attributes on the `a:path` element, respectively. The top, left
  corner of the path bounding box is (0, 0) and the bottom, right corner is
  at (`h`, `w`). Coordinates are positive integers in the range 0 to
  27,273,042,316,900 (about 2^44.63).


Resources
---------

* Office Open XML - Custom Geometry
  http://officeopenxml.com/drwSp-custGeom.php


XML Specimens
-------------

.. highlight:: xml

::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="7" name="Freeform 6"/>
      <p:cNvSpPr/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="5259090" y="708978"/>
        <a:ext cx="2719145" cy="1012826"/>
      </a:xfrm>
      <a:custGeom>
        <a:avLst/>
        <a:gdLst>
          <a:gd name="connsiteX0" fmla="*/ 0 w 2719145"/>
          <a:gd name="connsiteY0" fmla="*/ 0 h 1012826"/>
          <a:gd name="connsiteX1" fmla="*/ 498640 w 2719145"/>
          <a:gd name="connsiteY1" fmla="*/ 724560 h 1012826"/>
          <a:gd name="connsiteX2" fmla="*/ 1862108 w 2719145"/>
          <a:gd name="connsiteY2" fmla="*/ 1012826 h 1012826"/>
          <a:gd name="connsiteX3" fmla="*/ 2687980 w 2719145"/>
          <a:gd name="connsiteY3" fmla="*/ 599905 h 1012826"/>
          <a:gd name="connsiteX4" fmla="*/ 2719145 w 2719145"/>
          <a:gd name="connsiteY4" fmla="*/ 475249 h 1012826"/>
        </a:gdLst>
        <a:ahLst/>
        <a:cxnLst>
          <a:cxn ang="0">
            <a:pos x="connsiteX0" y="connsiteY0"/>
          </a:cxn>
          <a:cxn ang="0">
            <a:pos x="connsiteX1" y="connsiteY1"/>
          </a:cxn>
          <a:cxn ang="0">
            <a:pos x="connsiteX2" y="connsiteY2"/>
          </a:cxn>
          <a:cxn ang="0">
            <a:pos x="connsiteX3" y="connsiteY3"/>
          </a:cxn>
          <a:cxn ang="0">
            <a:pos x="connsiteX4" y="connsiteY4"/>
          </a:cxn>
        </a:cxnLst>
        <a:rect l="l" t="t" r="r" b="b"/>
        <a:pathLst>
          <a:path w="2719145" h="1012826">
            <a:moveTo>
              <a:pt x="0" y="0"/>
            </a:moveTo>
            <a:lnTo>
              <a:pt x="498640" y="724560"/>
            </a:lnTo>
            <a:lnTo>
              <a:pt x="1862108" y="1012826"/>
            </a:lnTo>
            <a:lnTo>
              <a:pt x="2687980" y="599905"/>
            </a:lnTo>
            <a:lnTo>
              <a:pt x="2719145" y="475249"/>
            </a:lnTo>
          </a:path>
        </a:pathLst>
      </a:custGeom>
      <a:ln w="50800">
        <a:solidFill>
          <a:schemeClr val="accent2"/>
        </a:solidFill>
        <a:prstDash val="sysDot"/>
      </a:ln>
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
    <p:txBody>
      <a:bodyPr rtlCol="0" anchor="ctr"/>
      <a:lstStyle/>
      <a:p>
        <a:pPr algn="ctr"/>
        <a:endParaRPr lang="en-US"/>
      </a:p>
    </p:txBody>
  </p:sp>


XML Schema excerpt
------------------

::

  <xsd:complexType name="CT_Shape">  <!-- p:sp element -->
    <xsd:sequence>
      <xsd:element name="nvSpPr" type="CT_ShapeNonVisual"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties"/>
      <xsd:element name="style"  type="a:CT_ShapeStyle"        minOccurs="0"/>
      <xsd:element name="txBody" type="a:CT_TextBody"          minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="useBgFill" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ShapeProperties">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="xfrm"                type="CT_Transform2D"            minOccurs="0"/>
      <xsd:choice minOccurs="0">  <!-- EG_Geometry -->
        <xsd:element name="custGeom" type="CT_CustomGeometry2D"/>
        <xsd:element name="prstGeom" type="CT_PresetGeometry2D"/>
      </xsd:choice>
      <xsd:group    ref="EG_FillProperties"                                    minOccurs="0"/>
      <xsd:element name="ln"                  type="CT_LineProperties"         minOccurs="0"/>
      <xsd:group    ref="EG_EffectProperties"                                  minOccurs="0"/>
      <xsd:element name="scene3d"             type="CT_Scene3D"                minOccurs="0"/>
      <xsd:element name="sp3d"                type="CT_Shape3D"                minOccurs="0"/>
      <xsd:element name="extLst"              type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
  </xsd:complexType>

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

  <xsd:complexType name="CT_Path2DList">
    <xsd:sequence>
      <xsd:element name="path" type="CT_Path2D" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Path2D">
    <xsd:choice minOccurs="0" maxOccurs="unbounded">
      <xsd:element name="close"      type="CT_Path2DClose"/>
      <xsd:element name="moveTo"     type="CT_Path2DMoveTo"/>
      <xsd:element name="lnTo"       type="CT_Path2DLineTo"/>
      <xsd:element name="arcTo"      type="CT_Path2DArcTo"/>
      <xsd:element name="quadBezTo"  type="CT_Path2DQuadBezierTo"/>
      <xsd:element name="cubicBezTo" type="CT_Path2DCubicBezierTo"/>
    </xsd:choice>
    <xsd:attribute name="w"           type="ST_PositiveCoordinate" default="0"/>
    <xsd:attribute name="h"           type="ST_PositiveCoordinate" default="0"/>
    <xsd:attribute name="fill"        type="ST_PathFillMode"       default="norm"/>
    <xsd:attribute name="stroke"      type="xsd:boolean"           default="true"/>
    <xsd:attribute name="extrusionOk" type="xsd:boolean"           default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Path2DMoveTo">
    <xsd:sequence>
      <xsd:element name="pt" type="CT_AdjPoint2D"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_AdjPoint2D">
    <xsd:attribute name="x" type="ST_AdjCoordinate" use="required"/>
    <xsd:attribute name="y" type="ST_AdjCoordinate" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_GeomGuideName">
    <xsd:restriction base="xsd:token"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_GeomGuideFormula">
    <xsd:restriction base="xsd:string"/>
  </xsd:simpleType>

  <xsd:complexType name="CT_GeomGuide">
    <xsd:attribute name="name" type="ST_GeomGuideName"    use="required"/>
    <xsd:attribute name="fmla" type="ST_GeomGuideFormula" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_GeomGuideList">
    <xsd:sequence>
      <xsd:element name="gd" type="CT_GeomGuide" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:simpleType name="ST_AdjCoordinate">
    <xsd:union memberTypes="ST_Coordinate ST_GeomGuideName"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_AdjCoordinate">
    <xsd:union memberTypes="ST_Coordinate ST_GeomGuideName"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_PositiveCoordinate">
    <xsd:restriction base="xsd:long">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="27273042316900"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_Coordinate">
    <xsd:union memberTypes="ST_CoordinateUnqualified s:ST_UniversalMeasure"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_CoordinateUnqualified">
    <xsd:restriction base="xsd:long">
      <xsd:minInclusive value="-27273042329600"/>
      <xsd:maxInclusive value="27273042316900"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_UniversalMeasure">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)"/>
    </xsd:restriction>
  </xsd:simpleType>
