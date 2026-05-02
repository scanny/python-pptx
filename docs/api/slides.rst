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


|Transition| objects
--------------------

The |Transition| object is accessed as the :attr:`~pptx.slide.Slide.transition`
property of |Slide|. It provides read/write access to the slide-transition
XML element. This is the Foundation F8 MVP surface; structured animation
effects (entrance/exit/emphasis/motion-path) and the MORPH
``mc:AlternateContent`` wrapper are tracked as downstream items — see
``docs/dev/analysis/f8-animations-transitions.rst``.

.. autoclass:: pptx.slide.Transition()
   :members:
   :undoc-members:


|SlideLayouts| objects
----------------------

The |SlideLayouts| object is accessed using the
:attr:`~pptx.slide.SlideMaster.slide_layouts` property of |SlideMaster|, typically::

    >>> from pptx import Presentation
    >>> prs = Presentation()
    >>> slide_layouts = prs.slide_master.slide_layouts

As a convenience, since most presentations have only a single slide master, the
|SlideLayouts| collection for the first master may be accessed directly from the
|Presentation| object::

    >>> slide_layouts = prs.slide_layouts

This class is not intended to be constructed directly.

.. autoclass:: pptx.slide.SlideLayouts()
   :members:
   :exclude-members: element, parent
   :inherited-members:
   :undoc-members:


|SlideLayout| objects
---------------------

.. autoclass:: pptx.slide.SlideLayout
   :members:
   :exclude-members: iter_cloneable_placeholders


|_HeaderFooter| objects
-----------------------

A ``_HeaderFooter`` object is returned by the ``header_footer`` property on a
|SlideMaster|, |SlideLayout|, or |NotesMaster|. It provides Boolean toggles for the
slide-number, header, footer, and date placeholders (the four attributes of the
``<p:hf>`` element). A value of ``False`` hides the corresponding placeholder; the
default for each toggle is ``True`` (placeholder visible) to match the XSD default.

This class is not intended to be constructed directly.

.. autoclass:: pptx.slide._HeaderFooter()
   :members:
   :undoc-members:


|SlideMasters| objects
----------------------

The |SlideMasters| object is accessed using the
:attr:`~pptx.presentation.slide_masters` property of |Presentation|, typically::

    >>> from pptx import Presentation
    >>> prs = Presentation()
    >>> slide_masters = prs.slide_masters

As a convenience, since most presentations have only a single slide master, the
first master may be accessed directly from the |Presentation| object without indexing
the collection::

    >>> slide_master = prs.slide_master

This class is not intended to be constructed directly.

.. autoclass:: pptx.slide.SlideMasters()
   :members:
   :exclude-members: element, parent
   :inherited-members:
   :undoc-members:

|SlideMaster| objects
---------------------

.. autoclass:: pptx.slide.SlideMaster
   :members:
   :exclude-members: related_slide_layout, sldLayoutIdLst


|SlidePlaceholders| objects
---------------------------

See :ref:`placeholder_api` for the placeholder collection on a slide.


|NotesSlide| objects
--------------------

.. autoclass:: pptx.slide.NotesSlide
   :members:
   :exclude-members: clone_master_placeholders
   :inherited-members:
