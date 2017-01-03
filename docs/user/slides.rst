
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

Right now, adding a slide is the only operation on the slide collection. On the
backlog at the time of writing is deleting a slide and moving a slide to
a different position in the list. Copying a slide from one presentation to
another turns out to be pretty hard to get right in the general case, so that
probably won't come until more of the backlog is burned down.


Up next ...
-----------

Ok, now that we have a new slide, let's talk about how to put something on
it ...
