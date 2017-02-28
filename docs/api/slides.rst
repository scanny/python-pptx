.. _slides_api:

Slides
======

|Slides| objects
-----------------

The |Slides| object is accessed using the
:attr:`~pptx.presentation.Presentation.slides` property of |Presentation|. It
is not intended to be constructed directly.

.. autoclass:: pptx.slide.Slides()
   :members:
   :member-order: bysource
   :undoc-members:


|Slide| objects
---------------

An individual |Slide| object is accessed by index from |Slides| or as the
return value of :meth:`add_slide`.

.. autoclass:: pptx.slide.Slide()
   :members:
   :exclude-members: part
   :inherited-members:
   :undoc-members:


|SlideLayout| objects
---------------------

.. autoclass:: pptx.slide.SlideLayout
   :members:
   :exclude-members: iter_cloneable_placeholders


|SlideMaster| objects
---------------------

.. autoclass:: pptx.slide.SlideMaster
   :members:
   :exclude-members: related_slide_layout, sldLayoutIdLst


|SlidePlaceholders| objects
---------------------------

.. autoclass:: pptx.shapes.shapetree.SlidePlaceholders
   :members:


|NotesSlide| objects
--------------------

.. autoclass:: pptx.slide.NotesSlide
   :members:
   :exclude-members: clone_master_placeholders
   :inherited-members:
