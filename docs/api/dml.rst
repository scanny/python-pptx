
.. _dml_api:

DrawingML objects
=================

Low-level drawing elements like fill and color that appear repeatedly in
various aspects of shapes.


.. currentmodule:: pptx.dml.fill


|FillFormat| objects
--------------------

.. autoclass:: FillFormat
   :members:
   :exclude-members: from_fill_parent
   :undoc-members:


.. currentmodule:: pptx.dml.line


|LineFormat| objects
--------------------

.. autoclass:: LineFormat
   :members:
   :undoc-members:


.. currentmodule:: pptx.dml.color


|ColorFormat| objects
---------------------

.. autoclass:: ColorFormat
   :members: brightness, rgb, theme_color, type
   :undoc-members:


|RGBColor| objects
------------------

.. autoclass:: RGBColor
   :members: from_string
   :undoc-members:
