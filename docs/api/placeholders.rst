.. _placeholder_api:

Placeholders
============

The following classes represent placeholder shapes. A placeholder most
commonly appears on a slide, but also appears on a slide layout and a slide
master. The role of a master placeholder and layout placeholder differs from
that of a slide placeholder and these roles are reflected in the distinct
classes for each.

There are a larger variety of slide placeholders to accomodate their more
complex and varied behaviors.


|SlidePlaceholders| objects
---------------------------

The placeholders collection on a slide is accessed as
``slide.placeholders``. It supports iteration, ``len()``, and
dictionary-style lookup by placeholder ``idx`` (not list position)::

    title = slide.placeholders[0]         # raises KeyError if absent
    body = slide.placeholders.get(1)      # returns None if absent

.. autoclass:: pptx.shapes.shapetree.SlidePlaceholders()
   :members: get
   :special-members: __getitem__, __iter__, __len__


|MasterPlaceholder| objects
---------------------------

.. autoclass:: pptx.shapes.placeholder.MasterPlaceholder()
   :members:
   :exclude-members:
      adjustments, get_or_add_ln, has_chart, has_table, idx, ln, orient,
      part, ph_type, shape_type, sz
   :inherited-members:
   :undoc-members:


|LayoutPlaceholder| objects
---------------------------

.. autoclass:: pptx.shapes.placeholder.LayoutPlaceholder()
   :members:
   :undoc-members:


ChartPlaceholder objects
------------------------

.. autoclass:: pptx.shapes.placeholder.ChartPlaceholder()
   :members:
   :exclude-members:
      has_chart, has_table, has_text_frame, part
   :inherited-members:
   :undoc-members:


PicturePlaceholder objects
--------------------------

.. autoclass:: pptx.shapes.placeholder.PicturePlaceholder()
   :members:
   :exclude-members:
      has_chart, has_table, has_text_frame, part
   :inherited-members:
   :undoc-members:


TablePlaceholder objects
------------------------

.. autoclass:: pptx.shapes.placeholder.TablePlaceholder()
   :members:
   :inherited-members:
   :exclude-members:
      has_chart, has_table, has_text_frame, is_placeholder,
      part
   :undoc-members:


PlaceholderGraphicFrame objects
-------------------------------

.. autoclass:: pptx.shapes.placeholder.PlaceholderGraphicFrame()
   :members:
   :inherited-members:
   :exclude-members:
      chart_part, has_text_frame, is_placeholder, part
   :undoc-members:


PlaceholderPicture objects
--------------------------

.. autoclass:: pptx.shapes.placeholder.PlaceholderPicture()
   :members:
   :inherited-members:
   :exclude-members:
      get_or_add_ln, has_chart, has_table, has_text_frame, ln,
      part
   :undoc-members:


_PlaceholderFormat objects
--------------------------

.. autoclass:: pptx.shapes.base._PlaceholderFormat()
   :members:
   :inherited-members:
   :undoc-members:
