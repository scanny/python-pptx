
.. _dml_api:

DrawingML objects
=================

Low-level drawing elements like fill and color that appear repeatedly in
various aspects of shapes.


.. currentmodule:: pptx.dml.fill


|FillFormat| objects
--------------------

.. A |Table| object is added to a slide using the
.. :meth:`~.ShapeCollection.add_table` method on |ShapeCollection|.

.. autoclass:: FillFormat
   :members: background, fore_color, solid, type


.. currentmodule:: pptx.dml.color


|ColorFormat| objects
---------------------

.. autoclass:: ColorFormat
   :members: brightness, rgb, theme_color, type


|RGBColor| objects
------------------

.. autoclass:: RGBColor
   :members: from_string
