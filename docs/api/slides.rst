.. _slides_api:

Slides
======

|Slides| objects
-----------------

The |Slides| object is accessed using the
:attr:`~pptx.presentation.Presentation.slides` property of |Presentation|. It
is not intended to be constructed directly.

.. autoclass:: pptx.slide.Slides()
   :members: add_slide
   :member-order: bysource
   :undoc-members:


|Slide| objects
---------------

|Slide| objects are accessed by index from ``prs.slides`` or as a return value
from ``slides.add_slide()``.

.. autoclass:: pptx.parts.slide.Slide()
   :members:
   :exclude-members:
      add_chart_part, after_unmarshal, before_marshal, blob, content_type,
      drop_rel, get_image, get_or_add_image_part, load, load_rel, new,
      package, part, part_related_by, relate_to, related_parts, rels,
      spTree, target_ref
   :inherited-members:
   :undoc-members:


|SlideLayout| objects
---------------------

.. autoclass:: pptx.parts.slidelayout.SlideLayoutPart
   :members:
   :exclude-members: iter_cloneable_placeholders


|SlideMaster| objects
---------------------

.. autoclass:: pptx.parts.slidemaster.SlideMasterPart
   :members:
   :exclude-members: sldLayoutIdLst


|SlidePlaceholders| objects
---------------------------

.. autoclass:: pptx.shapes.factory.SlidePlaceholders
   :members:
