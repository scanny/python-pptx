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
XML element, including the Office 2010 **MORPH** transition (assigned via
``slide.transition.type = PP_TRANSITION_TYPE.MORPH``) which python-pptx
writes wrapped in an ``mc:AlternateContent`` with a ``p:fade`` fallback,
mirroring the form PowerPoint itself emits. The matching granularity of
a MORPH transition is selected via :attr:`~pptx.slide.Transition.morph_option`
(one of ``"byObject"`` / ``"byWord"`` / ``"byChar"``).

Structured animation effects (entrance / exit / emphasis / motion-path)
are tracked as downstream items — see
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


Google Slides interop
~~~~~~~~~~~~~~~~~~~~~

Decks exported from Google Slides frequently strip the ``@type`` attribute
from each ``<p:sldLayout>`` element and leave ``p:cSld/@name`` empty, which
makes :meth:`SlideLayouts.get_by_name` return ``None`` for lookups like
``slide_layouts.get_by_name("Blank")`` (see issue #864). Two mitigations ship
with python-pptx:

* :attr:`SlideLayout.name` falls back to a positional
  ``"Layout N"`` (1-based index within the parent master's ``slide_layouts``)
  when the underlying attribute is empty, so layouts remain
  distinguishable in user-facing output.

* :meth:`SlideLayouts.get_by_type` looks each layout up by its
  ``ST_SlideLayoutType`` token (e.g. ``"title"``, ``"blank"``,
  ``"cust"``) rather than by name. :attr:`SlideLayout.slide_layout_type`
  exposes the underlying attribute directly and defaults to ``"cust"``
  when absent, matching the ECMA-376 Part 1 §19.3.1.39 default. Layouts
  that Google Slides preserves verbatim will still carry a useful
  ``@type`` token, and for layouts authored entirely in Google Slides
  (all ``"cust"``) a positional index remains the reliable identifier.


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


|_EffectiveBackground| objects
------------------------------

An ``_EffectiveBackground`` object is returned by the ``effective_background``
property on a |Slide|, |SlideLayout|, or |SlideMaster|. Unlike |_Background|,
reading through this proxy is side-effect free: no accessor materializes a
``p:bgPr/a:noFill`` subtree on the underlying XML. It resolves the
inheritance chain *slide → layout → master* and exposes the first ancestor
carrying an explicit ``p:bg`` child.

This addresses the read path reported in issue #809 — callers that only
want to *read* the background PowerPoint would render (including inherited
layout / master backgrounds) should prefer ``effective_background`` over
``background``, whose ``.fill`` accessor is destructive.

This class is not intended to be constructed directly.

.. autoclass:: pptx.slide._EffectiveBackground()
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
