.. _GroupShape:

Group Shape
===========

A *group shape* is a container for other shapes.

In the PowerPoint UI, a group shape may be selected and moved as a unit, such
that the contained shapes retain their relative position to one another.

Certain operations can also be applied to all the shapes in a group by
applying that operation to the group shape.


Scope
-----

::

    Given a SlideShapes object as shapes containing a group shape
     Then shapes[0] is a GroupShape object

    * BaseShapeFactory() update
    * add tests.shapes.test_group_shape

Group shape inherits `BaseShape` properties and behaviors:

* [ ] Apparently, a chart can be a member of a group, but a table cannot. Also
      `SmartArt` and placeholders can only appear at the top level of the slide
      shape tree.

* [ ] Consider whether existing tests for things like `.add_connector()` should
      be moved to `GroupShape` instead. I think they should be.

* [ ] Class `GroupShape` needs to override `._next_shape_id` and use parent
      version or something.

* [ ] Consider updating `BaseShape.shape_type` to raise an exception (or at
      least a warning.

* [ ] Should height be settable? What happens if you change it? Does the group
      automatically scale?

      Consider overriding then calling super after documenting any behavior
      unique to a group shape.

* [ ] Consider whether `GroupShapes` should be located in
      `pptx.shapes.shapetree` module.

* [ ] Consider adding mixin `PlaceholderCloner` to host `.clone_placeholder()`
      and perhaps `.ph_basename` and `._next_ph_name` that can be added to
      `SlideShapes` and `NotesSlideShapes`.

      Maybe `_BaseShapes.ph_basename` moves to `SlideShapes`.


Group shape also inherits from `SlideShapes`
--------------------------------------------

Or maybe it's better if `GroupShape` has a `.shapes` property.

Maybe separate out `_BaseSingleShape` (i.e. not a group shape) for things
like `.has_chart`, `.is_placeholder`, etc. But actually most of the
properties are legitimate, only one or two like click_action aren't, maybe
better just to override those with a property that raises an exception.

* [ ] A group shape has no click action.

Maybe an iter_all() method on `SlideShapes` that does a depth-first traversal
of the shape graph.

Possible Scope
--------------

* `group_shape = shapes.group(shape_lst)` returns a newly-created group shape
  containing each shape in `shape_lst`.


MS API
------

* `Shape.GroupItems` - corresponds to `GroupShape.shapes`


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
