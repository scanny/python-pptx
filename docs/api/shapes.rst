.. _shape_api:

Shapes
======


.. currentmodule:: pptx.shapes.shapetree


|ShapeCollection| objects
--------------------------

The |ShapeCollection| object is typically encountered as the
:attr:`~BaseSlide.shapes` member of |Slide|. It is not intended to be
constructed directly.

.. autoclass:: ShapeCollection
   :inherited-members:
   :members:
   :member-order: bysource
   :undoc-members:
   :show-inheritance:


.. currentmodule:: pptx.shapes.shape


``Shape`` objects
-----------------

The following properties and methods are common to all shapes.

.. autoclass:: BaseShape
   :members:
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
