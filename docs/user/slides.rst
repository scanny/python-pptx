
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
  to add the slide is on the slide collection, not the presentation. By
  default ``add_slide()`` appends the new slide to the end of the collection;
  pass ``index=N`` to insert it at a specific zero-based position instead::

      # insert the new slide at the front of the deck
      prs.slides.add_slide(slide_layout, index=0)

      # insert the new slide as the third slide
      prs.slides.add_slide(slide_layout, index=2)

  Index semantics match :meth:`Slides.move_slide`: a negative ``index`` counts
  from the end (``-1`` is the last position), and an ``index`` beyond the end
  is clamped to the last position rather than raising.


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


Changing a slide's layout
-------------------------

Assigning to :attr:`.Slide.slide_layout` re-points an existing slide at
a different slide layout. The new layout may belong to any slide master
in the same presentation, including a master different from the
previously referenced layout's master::

    from pptx import Presentation

    prs = Presentation("example.pptx")
    slide = prs.slides[0]

    # -- swap within the same master (e.g. from "Title Slide" to
    # -- "Title and Content")
    slide.slide_layout = prs.slide_layouts[1]

    # -- or re-point at a layout belonging to a different slide master
    slide.slide_layout = prs.slide_masters[1].slide_layouts[0]

Only the slide-to-layout relationship is re-pointed; the slide's own
shape tree (placeholders, text boxes, pictures, and so on) is not
rewritten. The slide continues to inherit theme colours, fonts, and
background from the new layout's master through the usual PowerPoint
inheritance chain.

.. note::

   The new layout must belong to the same presentation as the slide.
   To use a layout defined in a different presentation, first import
   it via :meth:`.SlideMaster.add_layout_from` and then assign the
   returned layout to ``slide.slide_layout``.


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


Searching and copying slides between presentations
---------------------------------------------------

A common workflow (issue #879) is "find a specific slide in deck A by some
human-readable criterion and graft it into deck B". |pp| does not ship a
``Slides.find_by_title()`` method because no single criterion fits every
deck -- callers typically match by title text, by a distinguishing word in
any shape, by layout name, by ``slide_id`` (when the id has been stashed
somewhere durable), or by the slide's internal ``p:cSld/@name``. A short
helper plus :meth:`.Slides.add_slide_from_external` covers the whole
recipe::

    from pptx import Presentation


    def find_slide_by_title(prs, title):
        """Return the first slide in *prs* whose title text equals *title*.

        The title placeholder is :attr:`.SlideShapes.title` (the placeholder
        with ``idx == 0``) when the slide's layout provides one; slides
        built from a *Blank* layout do not have a title and are skipped.
        Returns ``None`` when no match is found.
        """
        for slide in prs.slides:
            title_shape = slide.shapes.title
            if title_shape is None:
                continue
            if not title_shape.has_text_frame:
                continue
            if title_shape.text_frame.text == title:
                return slide
        return None


    source = Presentation("library-deck.pptx")
    target = Presentation("all-hands.pptx")

    wanted = find_slide_by_title(source, "Q3 Highlights")
    if wanted is not None:
        target.slides.add_slide_from_external(
            wanted, target.slide_layouts[5]
        )
        target.save("all-hands.pptx")

A few variations on the search predicate cover the common cases:

* **Match by any shape's text** (the reporter of issue #696 wanted to
  locate a slide by a phrase that could appear in any text-bearing shape,
  not just the title)::

      def find_slide_containing(prs, phrase):
          for slide in prs.slides:
              for shape in slide.shapes:
                  if shape.has_text_frame and phrase in shape.text_frame.text:
                      return slide
          return None

* **Match by stable** ``slide_id`` -- once the caller has stashed an id,
  the deck can be reordered freely and the lookup still works
  (:meth:`.Slides.get_by_slide_id`)::

      sid = source.slides[2].slide_id      # stash once, durably
      wanted = source.slides.get_by_slide_id(sid)

* **Match by layout name** -- useful when a deck uses a custom layout
  such as "Executive Summary" for exactly one slide::

      def find_slide_by_layout_name(prs, layout_name):
          for slide in prs.slides:
              if slide.slide_layout.name == layout_name:
                  return slide
          return None

Once the source slide is in hand, :meth:`.Slides.add_slide_from_external`
clones it into the target deck (see "Copying a slide from one
presentation to another" above for the full-fidelity semantics -- image
and chart and OLE-object parts are all materialised in the target
package). The source presentation is not mutated by the copy, so the
source deck keeps working as a read-only "library" the caller pulls
slides out of. When no single layout on the target fits every imported
slide, call :meth:`.Slides.add_slide_from_external` once per slide and
pass the matching ``target.slide_layouts[...]`` each time.


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


Importing a slide layout from another master
--------------------------------------------

:meth:`.SlideMaster.add_layout_from` clones a slide layout from any other
master -- whether that master belongs to a different presentation or is a
second master inside the same deck -- onto this master. This addresses
the "I want to apply a layout from deck A to a slide in deck B" workflow
tracked by issue #1028::

    from pptx import Presentation

    source = Presentation("branded-template.pptx")
    target = Presentation("quarterly-deck.pptx")

    # -- bring a custom layout from the branded template onto the target deck --
    source_layout = source.slide_masters[0].slide_layouts.get_by_name(
        "Executive Summary"
    )
    target_master = target.slide_masters[0]
    new_layout = target_master.add_layout_from(source_layout)

    # -- the newly-added layout is appended to the target master's layouts --
    assert new_layout is target_master.slide_layouts[-1]

    # -- and a slide can now inherit from it --
    slide = target.slides.add_slide(new_layout)
    target.save("quarterly-deck.pptx")

The clone carries over:

* the layout's ``p:cSld`` shape tree (placeholders, geometry, text content),
* image relationships baked into the layout (deduplicated against images
  already in the target package), and
* any external hyperlink relationships.

The cloned layout inherits its *theme* -- the color scheme, font scheme,
and effect scheme -- from the **destination master**, not from the
source's original master. This matches PowerPoint's behavior when a user
drags a layout between masters in *Slide Master* view: the layout's
structure moves, but it adopts the receiving master's palette and typography.

:meth:`.SlideMaster.add_layout_from` raises :class:`ValueError` when the
destination master already contains a layout with the same name. Rename
the source or destination layout first to distinguish them; once they
have distinct names :meth:`~.SlideLayouts.get_by_name` can tell them
apart.

The "use an existing deck as a template" workflow remains useful for
bulk content merges -- open the layout-owning deck as the base, merge
content-carrying decks with :meth:`.Presentation.merge`, and drop any
unwanted starter slides::

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


Adding a brand-new layout to an existing master
-----------------------------------------------

:meth:`.SlideMaster.add_layout` creates a fresh slide layout on an existing
master -- the same operation PowerPoint's *Slide Master* view exposes
through *Insert Layout*. It addresses the workflow tracked by issue #413:
you want a new layout kind on *this* deck's master without importing it
from somewhere else. ``add_layout`` is the same-master cousin of
:meth:`.SlideMaster.add_layout_from` (which imports a layout from a
*different* master)::

    from pptx import Presentation

    prs = Presentation()
    master = prs.slide_masters[0]

    # -- the simplest form: a truly empty custom layout with no placeholders --
    scoreboard = master.add_layout("Scoreboard")

    # -- a slide can now inherit from the new layout --
    slide = prs.slides.add_slide(scoreboard)
    prs.save("deck.pptx")

The ``name`` argument is required; it is persisted on the new layout's
``p:cSld/@name`` and is the value :meth:`~.SlideLayouts.get_by_name` and
the PowerPoint UI show. It must be unique within the master's layouts --
``add_layout`` raises :class:`ValueError` on a name collision.

Pass ``based_on`` to seed the new layout from another layout on *this*
master; the shape tree, placeholders, and geometry of the basis layout
are deep-cloned onto the new layout, then the new layout's name is set
to the caller-supplied ``name``. This is the quickest way to produce a
customized variant of an existing layout::

    # -- start from the "Title Slide" layout and adapt it --
    title_layout = master.slide_layouts.get_by_name("Title Slide")
    branded = master.add_layout("Title Slide - Branded", based_on=title_layout)

    # -- tweak the clone without touching the original --
    for ph in branded.placeholders:
        ph.left = ph.left + 500000   # shift every placeholder to the right

The default ``based_on=None`` path produces a minimal layout with
``@type="cust"``, no placeholders, and a ``p:clrMapOvr/a:masterClrMapping``
element so the color map inherits from the master; both forms relate the
new layout part to *this* master (so it picks up the master's theme,
fonts, and color scheme) and append a fresh ``p:sldLayoutId`` entry.

``based_on`` must belong to the same master ``add_layout`` is called on;
use :meth:`.SlideMaster.add_layout_from` to clone a layout from a
*different* master into this one.


Adding shared shapes to a layout or master
------------------------------------------

|SlideLayout| and |SlideMaster| both expose a ``shapes`` collection that
supports the same shape-authoring methods as |SlideShapes| —
:meth:`~.SlideShapes.add_shape`, :meth:`~.SlideShapes.add_picture`,
:meth:`~.SlideShapes.add_textbox`, :meth:`~.SlideShapes.add_connector`,
:meth:`~.SlideShapes.add_group_shape`, and
:meth:`~.SlideShapes.build_freeform`. Shapes appended to a layout or
master appear on every slide inheriting from it, which makes them the
right home for branding elements (a client logo, a confidentiality
banner, a running footer) that should show up on every slide without
having to edit each slide individually.

Authoring a shared text box on a layout (issue #1044)::

    from pptx import Presentation
    from pptx.util import Inches, Pt

    prs = Presentation()
    layout = prs.slide_layouts[0]

    tb = layout.shapes.add_textbox(Inches(0.3), Inches(6.7), Inches(9), Inches(0.3))
    tf = tb.text_frame
    tf.text = "CONFIDENTIAL — Draft"
    tf.paragraphs[0].runs[0].font.size = Pt(10)

    slide = prs.slides.add_slide(layout)
    prs.save("deck.pptx")  # every slide built from this layout sees the banner

The same idiom works on ``prs.slide_master.shapes`` when the shared
content should apply to *every* layout (and therefore every slide)
rather than just one layout. Individual slides may opt out of
master-level shared shapes by assigning
``slide.show_master_shapes = False``.


Adding a watermark
------------------

A "watermark" in PowerPoint is typically a semi-transparent picture or
a lightly-colored piece of text that shows through on every slide of
the deck (issue #793). PowerPoint itself doesn't have a dedicated
watermark feature — authors add the watermark to the *slide master*
so that every slide inheriting from that master displays it. The same
approach works in python-pptx using the shared-shapes machinery
covered in the previous section plus :attr:`.Picture.transparency`.

**Picture watermark**

Add the image to ``prs.slide_master.shapes`` and dial back its opacity
with :attr:`~pptx.shapes.picture.Picture.transparency` (a percentage in
``[0.0, 100.0]`` — ``0`` is fully opaque, ``100`` is invisible)::

    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    master = prs.slide_masters[0]

    # -- centered on a 10" x 7.5" slide, roomy enough to read through --
    pic = master.shapes.add_picture(
        "company-logo.png", Inches(3), Inches(2.25), Inches(4), Inches(3)
    )
    pic.transparency = 80.0                    # 80% transparent

    prs.slides.add_slide(prs.slide_layouts[1])
    prs.save("deck.pptx")                      # watermark shows on every slide

The watermark is authored once on the master; any slide created from
any layout under that master picks it up automatically, with no
per-slide bookkeeping.

**Text watermark**

When a simple "DRAFT", "CONFIDENTIAL", or "SAMPLE" text stamp is
enough, a large lightly-colored textbox on the master works the same
way::

    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    prs = Presentation()
    master = prs.slide_masters[0]

    tb = master.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(1.5))
    tf = tb.text_frame
    tf.text = "DRAFT"
    run = tf.paragraphs[0].runs[0]
    run.font.size = Pt(96)
    run.font.color.rgb = RGBColor(0xC0, 0xC0, 0xC0)   # light gray

    prs.save("deck.pptx")

Use :attr:`~pptx.dml.color.ColorFormat.alpha` on the run color for an
even subtler effect — the text is drawn at the chosen color with
per-pixel translucency, layered over the slide's own content.

**Excluding the watermark from a single slide**

A title slide or cover slide often shouldn't carry the watermark.
Because the watermark is a master shape, opting out on one slide is a
single assignment::

    title_slide = prs.slides[0]
    title_slide.show_master_shapes = False   # hides *all* master shapes

See :attr:`.Slide.show_master_shapes` above for the full semantics.


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


Editing the footer, slide-number, or date placeholder text
----------------------------------------------------------

A common question (issue #823) is "how do I change the small character /
string in the lower-left (or lower-right) corner of every slide?" That
text is almost always one of the *latent* footer, date, or slide-number
placeholders authored on the slide master (and optionally overridden on
specific layouts). Under PowerPoint conventions these three placeholders
live at ``idx == 10`` (date), ``idx == 11`` (footer), and ``idx == 12``
(slide number) on **layouts** — but their ``idx`` is different on the
master (typically 2, 3, 4 in insertion order). The reliable way to find
each one is by its ``placeholder_format.type`` enum value, not by index::

    from pptx import Presentation
    from pptx.enum.shapes import PP_PLACEHOLDER

    prs = Presentation("deck.pptx")

    def set_placeholder_text(container, ph_type, text):
        """Set the text of the first ``ph_type`` placeholder on *container*.

        *container* can be a |SlideMaster|, |SlideLayout|, or |Slide|.
        Returns the placeholder that was edited, or |None| if the
        container has no placeholder of that type.
        """
        for ph in container.placeholders:
            if ph.placeholder_format.type == ph_type:
                ph.text = text
                return ph
        return None

    master = prs.slide_master
    set_placeholder_text(master, PP_PLACEHOLDER.FOOTER, "Confidential – Q3 2026")

    # overriding a single layout (e.g. the "Title Slide" layout only)
    title_layout = master.slide_layouts.get_by_name("Title Slide")
    set_placeholder_text(title_layout, PP_PLACEHOLDER.FOOTER, "")

    prs.save("deck.pptx")

Slide-level vs layout-level vs master-level edits follow the usual
inheritance order: a layout-level footer text overrides the master, and
a slide-level one overrides the layout. Because :meth:`.Slides.add_slide`
does **not** clone the latent placeholders onto the slide, most callers
edit the text on the master or on the layout and rely on PowerPoint's
inheritance to display it everywhere.

Controlling visibility is a separate concern and is handled by
``header_footer`` (see the previous section); editing the text does
nothing if the ``<p:hf>`` toggle has hidden that placeholder.

If the "character in the lower-left corner" is actually a decorative
non-placeholder shape baked into the master (e.g. a logo bitmap or a
stylized watermark), iterate ``master.shapes`` instead and match by
``name`` or by inspecting ``shape.text_frame.text``::

    for shape in prs.slide_master.shapes:
        if shape.has_text_frame and "watermark" in shape.text_frame.text.lower():
            shape.text_frame.text = ""          # or edit runs individually

    prs.save("deck.pptx")


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

The reverse lookup — resolving a stored ``slide_id`` back to its
|Slide| — is :meth:`.Slides.get_by_slide_id`::

    >>> sid = prs.slides[0].slide_id        # stash the id somewhere durable
    >>> prs.slides.move_slide(prs.slides[0], 2)
    >>> prs.slides.get_by_slide_id(sid)     # still finds the same slide
    <pptx.slide.Slide object at 0x...>
    >>> prs.slides.get_by_slide_id(9999) is None
    True
    >>> prs.slides.get_by_slide_id(9999, default="fallback")
    'fallback'

The lookup walks ``p:sldIdLst/p:sldId`` and resolves the matching entry's
``r:id`` relationship to a slide part, so it is O(n) in the number of
slides but survives any amount of reordering. It is the explicit-named
counterpart to :meth:`.Slides.get` and the sibling of
:meth:`.SlideMaster.get_layout`.

**Background inheritance**::

    >>> slide.follow_master_background
    True
    >>> # apply a custom background — breaks master inheritance
    >>> slide.background.fill.solid()
    >>> slide.background.fill.fore_color.rgb = RGBColor(0xFF, 0xAA, 0x00)
    >>> slide.follow_master_background
    False
    >>> # revert to master inheritance — PowerPoint's "Reset Background" button
    >>> slide.follow_master_background()
    >>> slide.follow_master_background
    True

:attr:`.Slide.follow_master_background` is a dual bool-like / callable
attribute. Reading it returns |True| when the slide has no ``p:bg`` child
(so the background is inherited from its layout / master) and |False|
when the slide carries its own explicit background. *Calling* it
(``slide.follow_master_background()``) drops the slide's ``p:bg`` child
if one is present — the equivalent of PowerPoint's *Reset Background*
button. The call returns the slide itself, so it can be chained. It is
a no-op on a slide that already follows the master background.

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

**Inspecting and copying raw background XML**::

    >>> # read-only, side-effect-free access to the p:bg element
    >>> slide.background.bg_element
    None                                  # inherited background
    >>> other.background.bg_element.tag   # explicit background
    '{http://schemas.openxmlformats.org/presentationml/2006/main}bg'

    >>> # propagate one slide's background markup onto another
    >>> slide.copy_background_from(other)

The ``background.bg_element`` property on :attr:`.Slide.background`
returns the underlying ``p:bg`` element or |None| when the slide
inherits its background from the master or layout. Unlike
:attr:`.Slide.background.fill`, reading ``bg_element`` does *not*
materialize a ``p:bg`` subtree when none is present, so inheritance
stays intact. Use it for raw-XML inspection or for hand-rolled
manipulation outside the :class:`.FillFormat` abstraction.

:meth:`.Slide.copy_background_from` deep-copies the source slide's
``p:bg`` subtree onto this slide. Passing a source that inherits (has
no explicit ``p:bg``) removes any explicit background on this slide and
restores inheritance. The copy is purely XML-level: a ``p:bgRef`` style
reference keyed on the master theme is copied verbatim, so copying a
theme-bound background between presentations with different masters may
yield unresolved style references.

**Reading the rendered background — inheritance-aware**::

    >>> # the slide itself has no p:bg — background is inherited
    >>> slide.background.bg_element
    None
    >>> eff = slide.effective_background
    >>> eff.source
    'layout'                      # this slide inherits from its layout
    >>> eff.fill.fore_color.rgb
    RGBColor(0xFF, 0x00, 0x00)    # the layout's color

:attr:`.Slide.effective_background` walks the inheritance chain
*slide → layout → master* and returns a side-effect-free
:class:`.Slide._EffectiveBackground` view of the first ancestor carrying
an explicit ``p:bg``. Unlike :attr:`.Slide.background`, reading through
this proxy does **not** materialize a ``p:bgPr/a:noFill`` subtree on the
slide, so inheritance stays intact. Use it whenever you need to
**read** the background color (or other fill properties) PowerPoint
would actually render, including when the color comes from the layout
or master (issue #809).

The returned proxy exposes:

* :attr:`~pptx.slide._EffectiveBackground.source` — ``"slide"``,
  ``"layout"``, or ``"master"``, indicating which ancestor supplied the
  resolved background.
* :attr:`~pptx.slide._EffectiveBackground.owner` — the slide-like
  object (|Slide| / |SlideLayout| / |SlideMaster|) owning the ``p:bg``.
* :attr:`~pptx.slide._EffectiveBackground.bg_element` — the resolved
  ``p:bg`` lxml element.
* :attr:`~pptx.slide._EffectiveBackground.fill` — a
  :class:`.FillFormat` reading the ``p:bg/p:bgPr`` non-destructively,
  or |None| when the resolved ``p:bg`` wraps a theme-keyed ``p:bgRef``
  (a background-style reference has no directly-readable color).

:attr:`.Slide.effective_background` returns |None| only when neither
the slide nor its layout nor its master declares a background — a rare
case, since PowerPoint-authored decks almost always carry a master-level
``p:bg``. :attr:`.SlideLayout.effective_background` and
:attr:`.SlideMaster.effective_background` provide the analogous walk on
their respective slide-like objects.

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
