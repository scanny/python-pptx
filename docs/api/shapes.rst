.. _shape_api:

Shapes
======


.. currentmodule:: pptx.shapes.shapetree


|ShapeCollection| objects
--------------------------

The |ShapeCollection| object is encountered as the :attr:`~BaseSlide.shapes`
member of |Slide|.

.. autoclass:: ShapeCollection
   :members: add_picture, add_shape, add_table, add_textbox, index,
             placeholders, title


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
