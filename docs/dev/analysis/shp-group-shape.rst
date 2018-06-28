.. _GroupShape:


Group Shape
===========

A *group shape* is a container for other shapes.

In the PowerPoint UI, a group shape may be selected and moved as a unit, such
that the contained shapes retain their relative position to one another.

Certain operations can also be applied to all the shapes in a group by
applying that operation to the group shape.

The *grouping* notion is *recursive*; a group shape can itself contain one or
more group shapes. The "root" shape tree of a slide is itself a group,
although it has a few distinctive behaviors. The XML element for a slide's
shape tree (`p:spTree`) is the same XML type as that for a group shape
(`p:grpSp`), namely `CT_GroupShape`.


PowerPoint Behaviors
--------------------

* A group shape can have no click action.


XML Semantics
-------------

* `chOff` and `chExt` represent child offset and child extents, respectively.
  These are used if the group itself is transformed, in particular when it is
  scaled.


MS API
------

* `Shape.GroupItems` - corresponds to `GroupShape.shapes`


XML Specimens
-------------

.. highlight:: xml

::

  <p:grpSp>
    <p:nvGrpSpPr>
      <p:cNvPr id="5" name="Group 4"/>
      <p:cNvGrpSpPr/>
      <p:nvPr/>
    </p:nvGrpSpPr>
    <p:grpSpPr>
      <a:xfrm>
        <a:off x="3347864" y="2204864"/>
        <a:ext cx="3506688" cy="2930624"/>
        <a:chOff x="3347864" y="2204864"/>
        <a:chExt cx="3506688" cy="2930624"/>
      </a:xfrm>
    </p:grpSpPr>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="2" name="Rounded Rectangle 1"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="3347864" y="2204864"/>
          <a:ext cx="914400" cy="914400"/>
        </a:xfrm>
        <a:prstGeom prst="roundRect">
          <a:avLst/>
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
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="3" name="Oval 2"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="5940152" y="2708920"/>
          <a:ext cx="914400" cy="914400"/>
        </a:xfrm>
        <a:prstGeom prst="ellipse">
          <a:avLst/>
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
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="4" name="Isosceles Triangle 3"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="4355976" y="4221088"/>
          <a:ext cx="1060704" cy="914400"/>
        </a:xfrm>
        <a:prstGeom prst="triangle">
          <a:avLst/>
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
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </p:txBody>
    </p:sp>
  </p:grpSp>


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_GroupShape">
    <xsd:sequence>
      <xsd:element name="nvGrpSpPr"      type="CT_GroupShapeNonVisual"/>
      <xsd:element name="grpSpPr"        type="a:CT_GroupShapeProperties"/>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="sp"           type="CT_Shape"/>
        <xsd:element name="grpSp"        type="CT_GroupShape"/>
        <xsd:element name="graphicFrame" type="CT_GraphicalObjectFrame"/>
        <xsd:element name="cxnSp"        type="CT_Connector"/>
        <xsd:element name="pic"          type="CT_Picture"/>
        <xsd:element name="contentPart"  type="CT_Rel"/>
      </xsd:choice>
      <xsd:element name="extLst"         type="CT_ExtensionListModify" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GroupShapeNonVisual">
    <xsd:sequence>
      <xsd:element name="cNvPr"      type="a:CT_NonVisualDrawingProps"/>
      <xsd:element name="cNvGrpSpPr" type="a:CT_NonVisualGroupDrawingShapeProps"/>
      <xsd:element name="nvPr"       type="CT_ApplicationNonVisualDrawingProps"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GroupShapeProperties">
    <xsd:sequence>
      <xsd:element name="xfrm"    type="CT_GroupTransform2D"       minOccurs="0"/>
      <xsd:group   ref="EG_FillProperties"                         minOccurs="0"/>
      <xsd:group   ref="EG_EffectProperties"                       minOccurs="0"/>
      <xsd:element name="scene3d" type="CT_Scene3D"                minOccurs="0"/>
      <xsd:element name="extLst"  type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="ST_BlackWhiteMode"/>
  </xsd:complexType>

  <xsd:complexType name="CT_GroupTransform2D">
    <xsd:sequence>
      <xsd:element name="off"   type="CT_Point2D"        minOccurs="0"/>
      <xsd:element name="ext"   type="CT_PositiveSize2D" minOccurs="0"/>
      <xsd:element name="chOff" type="CT_Point2D"        minOccurs="0"/>
      <xsd:element name="chExt" type="CT_PositiveSize2D" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="rot"   type="ST_Angle"    default="0"/>
    <xsd:attribute name="flipH" type="xsd:boolean" default="false"/>
    <xsd:attribute name="flipV" type="xsd:boolean" default="false"/>
  </xsd:complexType>

  <xsd:group name="EG_EffectProperties">
    <xsd:choice>
      <xsd:element name="effectLst" type="CT_EffectList"/>
      <xsd:element name="effectDag" type="CT_EffectContainer"/>
    </xsd:choice>
  </xsd:group>

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
