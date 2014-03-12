.. _slides_api:

Slides
======

.. currentmodule:: pptx.parts.slides

|SlideCollection| objects
--------------------------

The |SlideCollection| object is typically encountered as the
:attr:`~pptx.Presentation.slides` member of |Presentation|. It is not intended
to be constructed directly.

.. autoclass:: SlideCollection
   :members: add_slide
   :member-order: bysource
   :undoc-members:


|Slide| objects
---------------

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


|SlideMaster| objects
---------------------

.. currentmodule:: pptx.parts.slidemaster

.. autoclass:: SlideMaster
   :members:
   :exclude-members: sldLayoutIdLst

.. currentmodule:: pptx.parts.slides

.. autoclass:: SlideLayout
   :members:
