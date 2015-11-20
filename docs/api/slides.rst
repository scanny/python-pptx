.. _slides_api:

Slides
======

.. currentmodule:: pptx.parts.presentation

|_Slides| objects
--------------------------

The |_Slides| object is encountered as the :attr:`~pptx.Presentation.slides`
member of |Presentation|. It is not intended to be constructed directly.

.. autoclass:: _Slides
   :members: add_slide
   :member-order: bysource
   :undoc-members:


|Slide| objects
---------------

.. currentmodule:: pptx.parts.slide

|Slide| objects are accessed by index from ``prs.slides`` or as a return value
from ``slides.add_slide()``.

.. autoclass:: Slide
   :members:
   :exclude-members:
      after_unmarshal, before_marshal, blob, content_type, drop_rel, load,
      load_rel, new, package, part, part_related_by, relate_to,
      related_parts, rels, target_ref
   :inherited-members:
   :undoc-members:


|SlideLayout| objects
---------------------

.. currentmodule:: pptx.parts.slidelayout

.. autoclass:: SlideLayout
   :members:
   :exclude-members: iter_cloneable_placeholders


|SlideMaster| objects
---------------------

.. currentmodule:: pptx.parts.slidemaster

.. autoclass:: SlideMaster
   :members:
   :exclude-members: sldLayoutIdLst


|_SlidePlaceholders| objects
----------------------------

.. currentmodule:: pptx.parts.slide

.. autoclass:: _SlidePlaceholders
   :members:
