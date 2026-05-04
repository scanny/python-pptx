.. _shape_api:

Shapes
======

The following classes provide access to the shapes that appear on a slide and
the collections that contain them.


|SlideShapes| objects
---------------------

The |SlideShapes| object is encountered as the :attr:`~BaseSlide.shapes`
property of |Slide|.

.. autoclass:: pptx.shapes.shapetree.SlideShapes()
   :members:
   :inherited-members:
   :exclude-members: clone_placeholder, clone_layout_placeholders,
                     ph_basename


|GroupShapes| objects
---------------------

The |GroupShapes| object is encountered as the :attr:`~GroupShape.shapes`
property of |GroupShape|.

.. autoclass:: pptx.shapes.shapetree.GroupShapes()
   :members:
   :inherited-members:
   :exclude-members: clone_placeholder, ph_basename


Shape objects in general
------------------------

The following properties and methods are common to all shapes.

.. autoclass:: pptx.shapes.base.BaseShape()
   :members:
   :exclude-members: part
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.shapes.base.ThemeStyleRefs()
   :members:
   :show-inheritance:


|Shape| objects (AutoShapes)
----------------------------

The following properties and methods are defined for AutoShapes, which
include text boxes and placeholders.

.. autoclass:: pptx.shapes.autoshape.Shape()
   :members:
   :exclude-members: get_or_add_ln, ln
   :member-order: bysource
   :undoc-members:


|AdjustmentCollection| objects
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

An AutoShape is distinctive in that it can have *adjustments*, represented in
the PowerPoint user interface as small yellow diamonds that each allow
a parameter of the shape, such as the angle of an arrowhead, to be adjusted.
The |AdjustmentCollection| object holds these adjustment values for an
AutoShape, each of which is an |Adjustment| instance.

The |AdjustmentCollection| instance for an AutoShape is accessed using the
``Shape.adjustments`` property (read-only).

.. autoclass:: pptx.shapes.autoshape.AdjustmentCollection
   :members:
   :member-order: bysource
   :undoc-members:


|Adjustment| objects
~~~~~~~~~~~~~~~~~~~~

.. autoclass:: pptx.shapes.autoshape.Adjustment
   :members:
   :member-order: bysource
   :undoc-members:


|Connector| objects
-------------------

The following properties and methods are defined for Connector shapes:

.. autoclass:: pptx.shapes.connector.Connector()
   :members:
   :member-order: bysource
   :exclude-members: get_or_add_ln, ln


|ConnectorAdjustmentCollection| objects
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

An elbow (bent) or curved connector has one or more *adjustments*, represented
in the PowerPoint user interface as small yellow diamonds that each allow the
position of a bend point to be moved. The |ConnectorAdjustmentCollection|
object holds these adjustment values for a connector and is accessed via the
``Connector.adjustments`` property (read-only reference; the individual values
are read/write).

Each value is a |float| nominally in the range 0.0 to 1.0, interpreted as a
proportion of the connector's width or height. A value of 0.5 places the bend
at the midpoint; values outside ``[0.0, 1.0]`` are valid and correspond to
bend points outside the connector's bounding box. Straight connectors (e.g.
``MSO_CONNECTOR.STRAIGHT``) and two-segment connectors have no adjustments, so
``len(connector.adjustments) == 0`` for those preset types.

.. autoclass:: pptx.shapes.connector.ConnectorAdjustmentCollection
   :members:
   :member-order: bysource
   :undoc-members:


|FreeformBuilder| objects
-------------------------

The following properties and methods are defined for FreeformBuilder objects.
A freeform builder is used to create a shape with custom geometry:

.. autoclass:: pptx.shapes.freeform.FreeformBuilder()
   :exclude-members: new, shape_offset_x, shape_offset_y
   :members:
   :member-order: bysource
   :undoc-members:


PathGeometry objects
--------------------

A ``PathGeometry`` is an ordered sequence of ``Path`` objects returned by
``Shape.path_geometry``. Each ``Path`` is an ordered sequence of
drawing-operation value objects (``MoveTo``, ``LineTo``, ``CubicBezierTo``,
``QuadBezierTo``, ``ArcTo``, ``Close``) whose coordinates are |Length|
instances expressed in shape-local EMU.

.. autoclass:: pptx.shapes.geometry.PathGeometry()
   :members:
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.shapes.geometry.Path()
   :members:
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.shapes.geometry.MoveTo()
   :members:
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.shapes.geometry.LineTo()
   :members:
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.shapes.geometry.CubicBezierTo()
   :members:
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.shapes.geometry.QuadBezierTo()
   :members:
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.shapes.geometry.ArcTo()
   :members:
   :member-order: bysource
   :undoc-members:

.. autoclass:: pptx.shapes.geometry.Close()
   :members:
   :member-order: bysource
   :undoc-members:


``Picture`` objects
-------------------

The following properties and methods are defined for picture shapes.

.. autoclass:: pptx.shapes.picture.Picture()
   :inherited-members:
   :members:
   :exclude-members: get_or_add_ln, ln
   :member-order: bysource


|Movie| objects
---------------

A |Movie| is a picture-shape variant that places a video (or audio clip
with poster-frame) on a slide. Movie shapes are added with
:meth:`SlideShapes.add_movie` and expose media-specific properties
alongside the shared |Picture| surface.

.. autoclass:: pptx.shapes.picture.Movie()
   :inherited-members:
   :members:
   :exclude-members: get_or_add_ln, ln
   :member-order: bysource


|GraphicFrame| objects
----------------------

The following properties and methods are defined for graphic frame shapes.
A graphic frame is the shape containing a table, chart, or smart art.

.. autoclass:: pptx.shapes.graphfrm.GraphicFrame()
   :show-inheritance:
   :members:
   :exclude-members:
       chart_part, has_text_frame, is_placeholder, part, placeholder_format,
       shape_type
   :inherited-members:


|SmartArt| objects
------------------

The |SmartArt| proxy provides read-only access to the four XML parts
that define a SmartArt graphic — *diagramData*, *diagramLayout*,
*diagramColors*, and *diagramQuickStyle*. It is obtained from
:attr:`GraphicFrame.smart_art` when the graphic frame contains a
SmartArt diagram. The MVP exposes the raw bytes of each part; a
structured editing API is not provided at this tier (see
``docs/dev/analysis/f9-smartart.rst`` for the downstream roadmap).

.. autoclass:: pptx.shapes.graphfrm.SmartArt()
   :members:
   :member-order: bysource
   :undoc-members:


|Model3D| objects
-----------------

The |Model3D| proxy provides read-only access to an embedded 3D model
(PowerPoint 365 "Insert > 3D Models") carried in a graphic frame. It is
obtained from :attr:`GraphicFrame.model_3d` when the graphic frame
contains a 3D model (``am3d:model3D`` under ``a:graphicData`` with the
Microsoft extension URI ``.../2016/12/model3D``). The MVP is
detection-and-passthrough: camera, lighting, and scene-graph authoring
are deferred — see ``docs/dev/analysis/model-3d.rst`` for the roadmap.

.. autoclass:: pptx.shapes.model3d.Model3D()
   :members:
   :member-order: bysource
   :undoc-members:


|GroupShape| objects
--------------------

The following properties and methods are defined for group shapes. A group
shape acts as a container for other shapes.

Note that:

* A group shape has no text frame and cannot have one.
* A group shape cannot have a click action, such as a hyperlink.

.. autoclass:: pptx.shapes.group.GroupShape()
   :show-inheritance:
   :members:
   :exclude-members:
       has_chart, has_table, is_placeholder, placeholder_format
   :inherited-members:
