.. _shape_api:

Shapes
======


.. currentmodule:: pptx.parts.slide


|_SlideShapeTree| objects
-------------------------

The |_SlideShapeTree| object is encountered as the :attr:`~BaseSlide.shapes`
property of |Slide|.

.. autoclass:: _SlideShapeTree
   :members:
   :exclude-members: clone_layout_placeholders


.. currentmodule:: pptx.shapes.shape


``Shape`` objects
-----------------

The following properties and methods are common to all shapes.

.. autoclass:: BaseShape
   :members: has_textframe, id, is_placeholder, name, shape_type, text,
             textframe
   :member-order: bysource
   :undoc-members:


.. currentmodule:: pptx.shapes.autoshape


The following properties and methods are defined for auto shapes and text boxes.

.. autoclass:: Shape
   :members:
   :member-order: bysource
   :undoc-members:


|Adjustment| objects
---------------------

.. autoclass:: Adjustment
   :members:
   :member-order: bysource
   :undoc-members:


|AdjustmentCollection| objects
-------------------------------

An |AdjustmentCollection| object reference is accessed using the
``Shape.adjustments`` property (read-only).

.. autoclass:: AdjustmentCollection
   :members:
   :member-order: bysource
   :undoc-members:
