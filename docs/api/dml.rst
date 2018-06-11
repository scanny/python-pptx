.. _dml_api:

DrawingML objects
=================

Low-level drawing elements like fill and color that appear repeatedly in
various aspects of shapes.


|ChartFormat| objects
---------------------

.. autoclass:: pptx.dml.chtfmt.ChartFormat
   :members:


|FillFormat| objects
--------------------

.. autoclass:: pptx.dml.fill.FillFormat
   :members:
   :exclude-members: from_fill_parent
   :undoc-members:


|LineFormat| objects
--------------------

.. autoclass:: pptx.dml.line.LineFormat
   :members:
   :undoc-members:


|ColorFormat| objects
---------------------

.. autoclass:: pptx.dml.color.ColorFormat
   :members: brightness, rgb, theme_color, type
   :undoc-members:


|RGBColor| objects
------------------

.. autoclass:: pptx.dml.color.RGBColor
   :members: from_string
   :undoc-members:


|ShadowFormat| objects
----------------------

.. autoclass:: pptx.dml.effect.ShadowFormat
   :members:
   :undoc-members:
