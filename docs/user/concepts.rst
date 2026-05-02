Concepts
========

|pp| is completely object-oriented, and in general any operations you perform
with it will be on an object. The root object for a presentation is
|Presentation|. API details are provided on the modules pages, but here are
some basics to get you started, especially some relationships you might find
surprising at first.

A presentation is loaded by constructing a new |Presentation| instance,
passing in the path to a presentation to be loaded::

    from pptx import Presentation

    path = 'slide-deck-foo.pptx'
    prs = Presentation(path)

|pp| also contains a default template, and if you construct a |Presentation|
instance without a path, a presentation based on that default template is
loaded. This can be handy when you want to get started quickly, and most of the
examples in this documentation use the default template.::

    # start with default presentation
    prs = Presentation()

Note that there is currently no distinction between templates and presentations
in |pp| as there is in the PowerPoint® client, there are only presentations. To
use a "template" for a presentation you simply create a presentation with all
the styles, logo, and layouts you want, delete all the slides (or leave some in
if it suits), and then load that as your starting place.


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
    slide_master = prs.slide_masters[0]
    # is equivalent to
    slide_master = prs.slide_master


Slide layouts
-------------

Another fact that might be surprising is that slide layouts belong to a slide
master, not directly to a presentation, so normally you have to access the
slide layouts via their slide master. Since this is subject to the same 99%
situation described above, the slide layouts belonging to the first slide
master can also be accessed directly from the presentation via syntactic
sugar::

    prs = Presentation()
    title_slide_layout = prs.slide_masters[0].slide_layouts[0]
    # is equivalent to:
    title_slide_layout = prs.slide_layouts[0]


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
    title_slide_layout = prs.slide_layouts[0]
    new_slide = prs.slides.add_slide(title_slide_layout)


Reaching past the API: raw XML access
-------------------------------------

PowerPoint's OOXML schema is vast and |pp| does not yet wrap every element or
attribute it defines. When you need to read or modify something the public
API does not surface, you can drop down to the underlying ``lxml`` element
tree without leaving the library. Every proxy object descending from
``ElementProxy`` (presentation, slide, section, chart, …) exposes a public
``element`` property that returns its backing ``lxml`` element, and every
shape provides the same via ``shape.element``::

    # The underlying <p:sld> element
    slide_elm = slide.element

    # The underlying <p:sp> / <p:pic> / … element
    shape_elm = shape.element

    # The underlying <c:chartSpace> element
    chartSpace = chart.element

You can mutate these elements in place; the changes are serialized when you
call ``prs.save(...)``. For XPath-based shape lookup on a slide, see
:meth:`Slide.find_shapes_by_xpath` (issue #224) — it evaluates an XPath
against the slide's ``p:spTree`` with the standard OOXML namespace prefixes
(``p``, ``a``, ``r``, ``mc``, ``p14``, …) pre-registered and returns matching
shape proxies::

    shapes = slide.find_shapes_by_xpath(
        './/p:sp[p:nvSpPr/p:cNvPr/@name="Title 1"]'
    )

Raw-XML access is supported but advanced; see :doc:`../dev/raw-xml-access` in
the contributor guide for guidance on when (and when not) to use it, and on
pairing it with the ``xmlchemy`` descriptor layer for typesafe access.
