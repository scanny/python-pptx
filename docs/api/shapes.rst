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


|FreeformBuilder| objects
-------------------------

The following properties and methods are defined for FreeformBuilder objects.
A freeform builder is used to create a shape with custom geometry:

.. autoclass:: pptx.shapes.freeform.FreeformBuilder()
   :exclude-members: new, shape_offset_x, shape_offset_y
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
