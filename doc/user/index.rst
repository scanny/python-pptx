==================
User Documentation
==================

|python-pptx| is a pure-Python library for manipulating Open XML presentation
files, perhaps more commonly referred to as PowerPoint files for PowerPoint
2007 and later versions.

.. toctree::
   :titlesonly:
   :maxdepth: 9

   modules/index
   examples



Use cases
=========

The use case that drove me to begin work on this library has been to automate
the building of slides that are tedious to compose by hand. As an example,
consider the task of composing a slide with an array of 10 headshot images of
folks in a particular department, with the person's name and title next to
their picture. After doing this a dozen times and struggling to get all the
alignment and sizes to the point where my attention to detail is satisfied,
well, my coding fingers got quite itchy.

However I believe a broader application will be server-side document
generation on non-Windows server platforms, Linux primarily I expect. In my
organization, I have found an apparently insatiable demand for PowerPoint
documents as a means of communication. Once one rises beyond the level of
project manager it seems the willingness to interpret text longer than a
bullet point atrophies quite rapidly and PowerPoint becomes an everyday
medium. I've imagined it might be pretty cool to be able to generate a
"presentation-ready" deck for a salesperson that includes a particular subset
of the product catalog they could generate with a few clicks to use in a sales
presentation, for example. As you come up with applications I'd love to hear
about them.


Getting Started
===============

|python-pptx| is completely object-oriented, and in general any operations you
perform with it will be on an object. The root object for a presentation is
|Presentation|. API details are provided on the modules pages, but here are
some basics to get you started, especially some relationships you might find
surprising at first.

A presentation is loaded by constructing a new |Presentation| instance,
passing in the path to a presentation to be loaded::

    from pptx import Presentation
    
    path = 'slide-deck-foo.pptx'
    prs = Presentation(path)

:mod:`python-pptx` also contains a default template, and if you construct a
|Presentation| instance without a path, a presentation based on that default
template is loaded. This can be handy when you want to get started quickly,
and most of the examples in this documentation use the default template.::

    # start with default presentation
    prs = Presentation()

Note that there is currently no distinction between templates and presentations in |python-pptx| as there is in the PowerPoint® client, there
are only presentations. To use a "template" for a presentation you simply
create a presentation with all the styles, logo, and layouts you want, delete
all the slides (or leave some in if it suits), and then load that as your
starting place.


Slide masters
-------------

A presentation has a list of slide masters and a list of slides. Let's start
with a discussion of the slide masters.

One fact some find surprising (I did) is that a presentation file can have
more than one slide master. It's quite uncommon in my experience to find
presentations that make use of this feature, but it's entirely supported. The
only time I've seen this happen is when slides from a "foreign" presentation
are pasted into another deck; if you want the formatting and backgrounds from
the other deck to be preserved on the pasted-in slides, the slide master and
its slide layouts need to come with. Consequently, the presentation needs to
maintain a list of slide masters, not just a single one, even though perhaps
99% of the time you only ever use the one. To make things a little easier for
the 99% situation, you can refer to the first slide master as though it were
the only one::

    prs = Presentation()
    slidemaster = prs.slidemasters[0]
    # is equivalent to
    slidemaster = prs.slidemaster


Slide layouts
-------------

Another fact that might be surprising is that slide layouts belong to a slide
master, not directly to a presentation, so normally you have to access the
slide layouts via their slide master. Since this is subject to the same 99%
situation described above, the slide layouts belonging to the first slide
master can also be accessed directly from the presentation via syntactic
sugar::

    prs = Presentation()
    title_slidelayout = prs.slidemasters[0].slidelayouts[0]
    # is equivalent to:
    title_slidelayout = prs.slidelayouts[0]


Slides
------

The slides in a presentation belong to the presentation object and are
accessed using the ``slides`` attribute::

    prs = Presentation(path)
    first_slide = prs.slides[0]


Adding a slide
^^^^^^^^^^^^^^

Adding a slide is accomplished by calling the :meth:`add_slide` method on the
:attr:`slides` attribute of the presentation. A slide layout must be passed
in to specify the layout the new slide should take on::

    prs = Presentation()
    title_slidelayout = prs.slidelayouts[0]
    new_slide = prs.slides.add_slide(title_slidelayout)


.. |python-pptx| replace:: :mod:`python-pptx`
.. |Presentation| replace:: :class:`Presentation`