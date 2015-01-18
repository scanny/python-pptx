.. _shape_api:

Shapes
======

The following classes provide access to the shapes that appear on a slide and
the collections that contain them.


|SlideShapeTree| objects
------------------------

The |SlideShapeTree| object is encountered as the :attr:`~BaseSlide.shapes`
property of |Slide|.

.. autoclass:: pptx.shapes.shapetree.SlideShapeTree()
   :members:
   :exclude-members: clone_layout_placeholders


Shape objects in general
------------------------

The following properties and methods are common to all shapes.

.. autoclass:: pptx.shapes.base.BaseShape()
   :members:
   :exclude-members: part
   :member-order: bysource
   :undoc-members:


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


``Picture`` objects
-------------------

The following properties and methods are defined for picture shapes.

.. autoclass:: pptx.shapes.picture.Picture()
   :members:
   :exclude-members: get_or_add_ln, ln
   :member-order: bysource
   :undoc-members:


|GraphicFrame| objects
----------------------

The following properties and methods are defined for graphic frame shapes.
A graphic frame is the shape containing a table, chart, or smart art.

.. autoclass:: pptx.shapes.graphfrm.GraphicFrame()
   :show-inheritance:
   :members:
   :exclude-members:
       chart_part, has_text_frame, has_textframe, is_placeholder, part,
       placeholder_format, shape_type
   :inherited-members:
