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


Gradient stops
--------------

.. automethod:: pptx.dml.fill._GradientStops.add_stop
.. automethod:: pptx.dml.fill._GradientStops.remove


|LineFormat| objects
--------------------

.. autoclass:: pptx.dml.line.LineFormat
   :members:
   :undoc-members:


|LineEndFormat| objects
-----------------------

.. autoclass:: pptx.dml.line.LineEndFormat
   :members:
   :undoc-members:


|ColorFormat| objects
---------------------

.. autoclass:: pptx.dml.color.ColorFormat
   :members: alpha, brightness, rgb, theme_color, to_rgb, type
   :undoc-members:


|RGBColor| objects
------------------

.. autoclass:: pptx.dml.color.RGBColor
   :members: from_string
   :undoc-members:


|EffectFormat| objects
----------------------

.. autoclass:: pptx.dml.effect.EffectFormat
   :members:
   :undoc-members:


|ShadowFormat| objects
----------------------

.. autoclass:: pptx.dml.effect.ShadowFormat
   :members:
   :undoc-members:


|GlowFormat| objects
--------------------

.. autoclass:: pptx.dml.effect.GlowFormat
   :members:
   :undoc-members:


|ReflectionFormat| objects
--------------------------

.. autoclass:: pptx.dml.effect.ReflectionFormat
   :members:
   :undoc-members:


|SoftEdgeFormat| objects
------------------------

.. autoclass:: pptx.dml.effect.SoftEdgeFormat
   :members:
   :undoc-members:
