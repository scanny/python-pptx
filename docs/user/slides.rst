
Working with Slides
===================

Every slide in a presentation is based on a slide layout. Not surprising then
that you have to specify which slide layout to use when you create a new slide.
Let's take a minute to understand a few things about slide layouts that we'll
need so the slide we add looks the way we want it to.


Slide layout basics
-------------------

A slide layout is like a template for a slide. Whatever is on the slide layout
"shows through" on a slide created with it and formatting choices made on the
slide layout are inherited by the slide. This is an important feature for
getting a professional-looking presentation deck, where all the slides are
formatted consistently. Each slide layout is based on the slide master in
a similar way, so you can make presentation-wide formatting decisions on the
slide master and layout-specific decisions on the slide layouts. There can
actually be multiple slide masters, but I'll pretend for now there's only one.
Usually there is.

The presentation themes that come with PowerPoint have about nine slide
layouts, with names like *Title*, *Title and Content*, *Title Only*, and
*Blank*. Each has zero or more placeholders (mostly not zero), preformatted
areas into which you can place a title, multi-level bullets, an image, etc.
More on those later.

The slide layouts in a standard PowerPoint theme always occur in the same
sequence. This allows content from one deck to be pasted into another and be
connected with the right new slide layout:

* Title (presentation title slide)
* Title and Content
* Section Header (sometimes called Segue)
* Two Content (side by side bullet textboxes)
* Comparison (same but additional title for each side by side content box)
* Title Only
* Blank
* Content with Caption
* Picture with Caption

In |pp|, these are ``prs.slide_layouts[0]`` through ``prs.slide_layouts[8]``.
However, there's no rule they have to appear in this order, it's just
a convention followed by the themes provided with PowerPoint. If the deck
you're using as your template has different slide layouts or has them in
a different order, you'll have to work out the slide layout indices for
yourself. It's pretty easy. Just open it up in Slide Master view in PowerPoint
and count down from the top, starting at zero.

Now we can get to creating a new slide.


Adding a slide
--------------

Let's use the Title and Content slide layout; a lot of slides do::

    SLD_LAYOUT_TITLE_AND_CONTENT = 1

    prs = Presentation()
    slide_layout = prs.slide_layouts[SLD_LAYOUT_TITLE_AND_CONTENT]
    slide = prs.slides.add_slide(slide_layout)

A few things to note:

* Using a "constant" value like ``SLD_LAYOUT_TITLE_AND_CONTENT`` is up to you.
  If you're creating many slides it can be handy to have constants defined so
  a reader can more easily make sense of what you're doing. There isn't a set
  of these built into the package because they can't be assured to be right for
  the starting deck you're using.

* ``prs.slide_layouts`` is the collection of slide layouts contained in the
  presentation and has list semantics, at least for item access which is about
  all you can do with that collection at the moment. Using ``prs`` for the
  Presentation instance is purely conventional, but I like it and use it
  consistently.

  When the layout ordering in the template you're working from may change
  across revisions, a position-based index like ``slide_layouts[1]`` is
  brittle. Each layout carries a presentation-stable integer identifier
  (``p:sldLayoutId/@id``) that PowerPoint preserves when layouts are
  reordered, and :meth:`.SlideMaster.get_layout` looks a layout up by that
  id::

      # -- capture the id of the layout once, when the deck is authored --
      TITLE_AND_CONTENT_ID = prs.slide_master.slide_layouts[1].slide_layout_id

      # -- later, the deck may have been reorganised but the id is stable --
      layout = prs.slide_master.get_layout(TITLE_AND_CONTENT_ID)

  :meth:`.SlideLayouts.get_by_id` is the equivalent lookup on the
  :class:`~pptx.slide.SlideLayouts` collection
  (``prs.slide_master.slide_layouts.get_by_id(id)``). Both accept a
  ``default=`` keyword that is returned when no layout has the requested
  id — ``None`` by default, matching ``dict.get`` semantics.

* ``prs.slides`` is the collection of slides in the presentation, also has
  list semantics for item access, and len() works on it. Note that the method
  to add the slide is on the slide collection, not the presentation. The
  ``add_slide()`` method appends the new slide to the end of the collection. At
  the time of writing it's the only way to add a slide, but sooner or later
  I expect someone will want to insert one in the middle, and when they post
  a feature request for that I expect I'll add an ``insert_slide(idx, ...)``
  method.


Doing other things with slides
------------------------------

In addition to adding a slide, the slide collection supports reordering a slide
to a different position via :meth:`~pptx.slide.Slides.move_slide`. The slide's
``slide_id`` is unchanged by the move, so any references stored by id keep
pointing to the same slide::

    prs = Presentation("example.pptx")
    slide = prs.slides[0]
    prs.slides.move_slide(slide, 2)  # move to zero-based position 2

A ``new_idx`` beyond the end of the collection moves the slide to the last
position, and negative values count from the end (``-1`` is the last slot).

The slide collection also supports deleting a slide with
:meth:`~pptx.slide.Slides.delete`::

    prs = Presentation("example.pptx")
    prs.slides.delete(prs.slides[1])  # remove the second slide

:meth:`~pptx.slide.Slides.delete` removes the slide's entry from the
``p:sldIdLst`` and drops the presentation-part relationship to the slide part.
On save, the slide part itself and any image, media, or chart parts that were
only referenced from that slide are omitted from the saved package (they become
unreachable from the package root). Parts that are also referenced by another
surviving slide (for example a chart or image used on multiple slides) are
preserved.

After calling :meth:`~pptx.slide.Slides.delete`, the deleted |Slide| object
should not be used; most operations on it will raise an exception.


Duplicating a slide
-------------------

:meth:`.Slides.duplicate` creates a deep copy of an existing slide in the same
presentation::

    prs = Presentation("quarterly-deck.pptx")
    source = prs.slides[0]
    dup = prs.slides.duplicate(source)           # appends a copy at the end
    prs.slides.duplicate(source, index=1)        # or insert at a specific spot

The duplicate:

* inherits from the same slide layout as ``source``,
* contains a deep copy of the shape tree (text, shapes, tables, etc.), and
* shares the underlying image, chart, OLE-object, media, and hyperlink parts
  with ``source`` — they are *reused* rather than re-embedded, so duplicating
  a slide doesn't bloat the package.

The duplicate receives a freshly-allocated ``slide_id``; the source slide's
``slide_id`` is unchanged.

Notes attached to the source slide are *not* copied onto the duplicate,
because each notes slide carries a back-reference to its owning slide and
can't be shared. Accessing ``duplicate.notes_slide`` creates a fresh empty
notes slide on demand.


Copying a slide from one presentation to another
------------------------------------------------

:meth:`.Slides.add_slide_from_external` appends a clone of a slide from one
presentation to another. The caller must supply a target slide-layout that
belongs to the destination presentation, since cross-presentation layout
binding is what keeps the cloned slide formatted consistently::

    from pptx import Presentation

    source = Presentation("quarterly-deck.pptx")
    target = Presentation("all-hands.pptx")
    cloned = target.slides.add_slide_from_external(
        source.slides[3], target.slide_layouts[5]
    )
    target.save("all-hands.pptx")

The clone is *full-fidelity*: the cloned slide's shape tree, image and media
parts, charts (each receiving a distinct embedded workbook so PowerPoint's
"Edit Data" keeps working on both the original and the copy), embedded OLE
objects, and external hyperlinks are all materialised in the target
presentation's package. The notes-slide relationship is dropped -- a notes
slide carries a back-reference to its owning slide and cannot be shared.


Merging every slide of another presentation
-------------------------------------------

To append *every* slide of another presentation in one call use
:meth:`Presentation.merge`::

    from pptx import Presentation

    target = Presentation("all-hands.pptx")
    source = Presentation("quarterly-deck.pptx")

    appended = target.merge(source)          # returns the list of new slides
    target.save("all-hands.pptx")

Each cloned slide is bound to the layout at the *same index* in the target's
primary slide master as its source's layout occupied in the source master
(with the last layout used as a fallback when the target master has fewer
layouts). Callers who need strict per-slide layout control can reassign
``slide.slide_layout`` after the merge, or bypass :meth:`merge` entirely and
drive :meth:`Slides.add_slide_from_external` per-slide.


Importing slide layouts from another presentation
-------------------------------------------------

python-pptx does not offer a direct "copy this layout into my deck" API
because importing a layout from deck A into deck B requires resolving its
slide-master, theme, and color / font / format scheme relationships into
B's package -- a deeper-reaching cross-package clone than the shipped
slide-copy pipeline.

The practical workaround is to open the presentation that *owns the
desired layouts* as the base, merge content-carrying decks into it, and
delete any unwanted starter slides::

    from pptx import Presentation

    # -- open the layout source as the base; the result inherits its
    # -- slide_layouts verbatim.
    base = Presentation("branded-template.pptx")

    # -- starter-slide indices to prune after the merge --
    starter_count = len(base.slides)

    # -- merge content from other decks; merged slides bind to base's
    # -- layouts by index (see Presentation.merge docstring).
    base.merge(Presentation("quarterly-deck.pptx"))
    base.merge(Presentation("all-hands.pptx"))

    # -- drop the starter slides, keeping only the merged content --
    for slide in list(base.slides)[:starter_count]:
        base.slides.delete(slide)

    base.save("combined.pptx")

The layouts of ``branded-template.pptx`` survive the whole procedure --
:meth:`Slides.delete` only removes slides, not the layouts they referenced,
so the target keeps its layout list intact while gaining the content of
every merged deck.


Header, footer, slide number, and date placeholders
---------------------------------------------------

Slide masters, slide layouts, and the notes master carry a small `<p:hf>`
element with four Boolean attributes controlling whether the header,
footer, slide-number, and date placeholders are visible on descendant
slides. python-pptx exposes each attribute through a ``header_footer`` property::

    >>> prs = Presentation()
    >>> master = prs.slide_master
    >>> master.header_footer.slide_number_visible
    True
    >>> master.header_footer.slide_number_visible = False       # hide slide numbers
    >>> master.header_footer.footer_visible = False             # hide footers too

The same attribute exists on every |SlideLayout|, letting you override a
master-level choice for an individual layout. PowerPoint treats an absent
``<p:hf>`` element as "all placeholders visible" — python-pptx keeps the XML
minimal and only writes a ``<p:hf>`` when at least one toggle is ``False``.

To display the *current* slide number inside a text frame, add an
auto-refresh field to a paragraph::

    >>> tf = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(1), Inches(0.5)).text_frame
    >>> field = tf.paragraphs[0].add_field("slidenum", "#")
    >>> field.field_type
    'slidenum'

Use ``"datetime"``, ``"datetimeFigureOut"``, or any of ``"datetime1"`` …
``"datetime13"`` for automatically-refreshing date fields, and ``"footer"`` for a
footer field. The ``text`` argument is the placeholder string PowerPoint shows
until the field is refreshed.


Reading slide properties
------------------------

Once you have a |Slide| object — whether freshly added with
:meth:`Slides.add_slide` or pulled from ``prs.slides[n]`` when opening an
existing deck — a handful of properties let you inspect the slide without
walking its XML.

**Slide identity**::

    >>> prs = Presentation("quarterly-deck.pptx")
    >>> slide = prs.slides[0]
    >>> slide.slide_id             # stable across slide reordering
    256
    >>> slide.name                 # the internal <p:cSld> @name (often empty)
    ''
    >>> slide.is_hidden
    False

``slide.slide_id`` is the stable, presentation-wide id PowerPoint assigns
when the slide is created; unlike the list index it does not shift when
slides are added, moved, or deleted. ``slide.name`` is the optional
internal name on ``<p:cSld>`` — PowerPoint leaves it empty for most
slides, so for human-readable identification you usually have to fall
back to the slide layout or the title text (see below).

**Hiding background graphics from the master**::

    >>> slide.show_master_shapes
    True
    >>> # hide the master's decorative shapes (e.g. a company logo) on this slide
    >>> slide.show_master_shapes = False

:attr:`.Slide.show_master_shapes` is the read/write boolean backing
PowerPoint's *Hide Background Graphics* checkbox (under *Design >
Format Background*). Assigning |False| writes ``p:sld/@showMasterSp="0"``
on the slide so the non-placeholder shapes baked into the master (a
company logo, a footer bar, and so on) no longer render underneath
this slide's content. Master *placeholders* continue to inherit
normally; only decorative master shapes are affected. Assigning |True|
(the schema default) removes the attribute, restoring the default
master-shapes-visible behavior.

**The slide layout**::

    >>> layout = slide.slide_layout
    >>> layout.name
    'Title and Content'
    >>> layout.slide_master.name
    'Office Theme'

``slide.slide_layout`` returns the |SlideLayout| the slide was created
from. Its ``name`` is the display name PowerPoint shows in *Slide Master*
view ("Title Slide", "Title and Content", "Two Content", …) and is
typically the most reliable way to classify a slide by kind.

**The title placeholder**::

    >>> title_shape = slide.shapes.title
    >>> title_shape is None
    False
    >>> title_shape.text_frame.text
    'Q3 Highlights'
    >>> # equivalent short form for read/write access to the title string
    >>> title_shape.text
    'Q3 Highlights'

:attr:`.SlideShapes.title` returns the title placeholder, or |None| if
the slide's layout doesn't have one (for example, a *Blank* layout).
Under the hood this is just the placeholder with ``idx == 0``, which by
convention is always the title slot.

**Iterating placeholders by idx**::

    >>> for ph in slide.placeholders:
    ...     print(ph.placeholder_format.idx, ph.placeholder_format.type, ph.name)
    0 TITLE (1) Title 1
    1 SUBTITLE (4) Subtitle 2

:attr:`.Slide.placeholders` yields placeholders in ``idx`` order. Each
placeholder exposes a :attr:`~._InheritsPlaceholderFormat.placeholder_format`
descriptor whose ``idx`` is the numeric slot and whose ``type`` is a
:ref:`PpPlaceholderType` enum member (``TITLE``, ``SUBTITLE``, ``BODY``,
``PICTURE``, …). A direct ``idx`` lookup works too::

    >>> subtitle = slide.placeholders[1]
    >>> subtitle.text_frame.text
    'Fiscal year review'

Note that ``slide.placeholders[idx]`` is keyed by placeholder ``idx``,
**not** by list position — so ``slide.placeholders[1]`` returns the
placeholder whose ``ph/@idx`` is ``1`` (the subtitle on a Title layout),
which may or may not be the second placeholder in iteration order.

**A convenience for non-title slide "names"**

Because ``slide.name`` is nearly always empty, code that needs to present
a human-readable label for a slide typically falls back to the title text,
or to the layout name::

    def slide_label(slide):
        """Best-effort human-readable identifier for *slide*."""
        if slide.name:
            return slide.name
        title = slide.shapes.title
        if title is not None and title.has_text_frame and title.text_frame.text:
            return title.text_frame.text
        return slide.slide_layout.name

This is enough to build a table-of-contents, navigate to a slide by
title, or filter slides by layout kind without leaving the
python-pptx public API.


Up next ...
-----------

Ok, now that we have a new slide, let's talk about how to put something on
it ...
