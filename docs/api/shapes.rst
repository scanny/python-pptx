.. _shape_api:

Shapes
======

The following classes provide access to the shapes that appear on a slide and
the collections that contain them.


|_SlideShapeTree| objects
-------------------------

The |_SlideShapeTree| object is encountered as the :attr:`~BaseSlide.shapes`
property of |Slide|.

.. autoclass:: pptx.parts.slide._SlideShapeTree()
   :members:
   :exclude-members: clone_layout_placeholders


Shape objects in general
------------------------

The following properties and methods are common to all shapes.

.. autoclass:: pptx.shapes.shape.BaseShape()
   :members:
   :exclude-members: part
   :member-order: bysource
   :undoc-members:


|Shape| objects (AutoShapes)
----------------------------

The following properties and methods are defined for AutoShapes, which
includes text boxes and placeholders.

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


``Placeholder`` objects
-----------------------

The following properties and methods are defined for placeholder shapes.

.. autoclass:: pptx.shapes.placeholder.BasePlaceholder()
   :members:
   :undoc-members:

.. autoclass:: pptx.shapes.placeholder.SlidePlaceholder()
   :members:
   :undoc-members:

.. autoclass:: pptx.shapes.placeholder.LayoutPlaceholder()
   :members:
   :undoc-members:

.. autoclass:: pptx.shapes.placeholder.MasterPlaceholder()
   :members:
   :exclude-members:
      adjustments, get_or_add_ln, ln, part, shape_type
   :inherited-members:
   :undoc-members:
