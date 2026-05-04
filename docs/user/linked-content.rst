.. _linked-content:

Linked / cross-referenced text
==============================

A recurring question (`issue #367
<https://github.com/scanny/python-pptx/issues/367>`_) is *"can I create a
link between two text fields so that editing one updates the other?"* —
essentially the PowerPoint equivalent of an Excel cell reference (``=A1``)
or a Word cross-reference field pointing at a bookmark on another page.
The short answer is **no**: PowerPoint has no native mechanism for
cross-slide, user-defined text linking, so python-pptx has no API to
author one.

This page explains what PowerPoint *does* offer in the field space, why
user-defined cross-references are not in that set, and the canonical
python-pptx recipe for keeping text in sync across slides —
re-authoring from a single source of truth in your generation script.


What PowerPoint does support
----------------------------

PowerPoint's ``<a:fld>`` element is a fixed, closed set of
*auto-refreshing* fields keyed off properties of the presentation or the
current slide. The field types PowerPoint writes — and the only ones
|pp| authors via :meth:`._Paragraph.add_field` — are:

* ``"slidenum"`` — the current slide's 1-based index (subject to
  :attr:`.Presentation.first_slide_num`).
* ``"datetime"`` / ``"datetimeFigureOut"`` — a date/time refreshed when
  the deck opens, rendered in the user's locale.
* ``"datetime1"`` … ``"datetime13"`` — the thirteen formatted date
  variants (``M/d/yyyy``, ``dddd, MMMM dd, yyyy``, etc.) that PowerPoint
  offers in the *Insert > Date & Time* dialog.
* ``"footer"`` — the footer string configured via the *Insert > Header &
  Footer* dialog (``<p:hf>`` visibility + the footer placeholder text on
  the master / layout — see the *Header, footer, slide number, and date
  placeholders* section of :doc:`slides`).

That is the full list. Each of these fields has its value computed by
PowerPoint (or its clone) at display time — the caller does not supply
the string, and there is no mechanism to point a field at an *arbitrary*
text run on another slide. OOXML has no "live reference to that text
frame over there" grammar analogous to Excel's cell reference.

A few adjacent features sometimes confused with cross-slide field links
are worth explicitly ruling out:

* **Excel-style cell references.** Excel's ``=Sheet1!A1`` lives inside
  the SpreadsheetML formula engine. DrawingML (the markup that backs
  PowerPoint shapes) has no formula layer.
* **Word-style cross-references.** Word's ``REF`` / ``PAGEREF`` fields
  target named bookmarks in the same document via WordprocessingML's
  ``<w:bookmarkStart>`` / ``<w:bookmarkEnd>`` grammar. PresentationML has
  no bookmark primitive.
* **Hyperlinks to another slide.** You *can* author a ``<a:hlinkClick>``
  whose target is another slide in the same deck (see
  :doc:`autoshapes` and :class:`.ActionSetting`), but this is a
  *navigation* link — clicking it jumps to the target slide — not a
  *content* link. It does not make the source text display the target
  text.
* **Linked Excel / OLE embeds.** Embedding a live-linked Excel range
  *does* give you a chunk of a PowerPoint slide that re-reads from an
  external ``.xlsx`` every time the deck opens; this is the only
  workflow where "the text updates when the source updates" is native to
  PowerPoint. See :doc:`ole-objects` for the embedding API and the
  :meth:`.SlideShapes.add_ole_object` entry point. The cost is
  substantial (the embed is an OLE object, not a run of text — it
  carries Excel-flavoured formatting, re-renders through Excel's
  rasterizer, and behaves differently from neighbouring shape text), so
  this is rarely the right answer for a plain "project name across every
  slide" use case.


The python-pptx recipe: one source of truth in your script
----------------------------------------------------------

The pragmatic solution — and the one every reporter on issue #367 has
ultimately adopted — is to treat your generation script as the single
source of truth, and re-author the same string onto every target slide
from a single variable. python-pptx is well suited to this because
round-trip fidelity means you can open an existing deck, walk the
slides, rewrite the relevant text, and save — without disturbing
anything else.

Example: stamp a project-name header on every slide::

    from pptx import Presentation
    from pptx.util import Inches, Pt

    PROJECT_NAME = "Apollo — Q3 2026 review"   # single source of truth

    prs = Presentation("template.pptx")

    for slide in prs.slides:
        # author the header as a new textbox in a fixed position
        tb = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9), Inches(0.4))
        p = tb.text_frame.paragraphs[0]
        run = p.add_run(PROJECT_NAME, bold=True, size=Pt(14))

    prs.save("out.pptx")

If the header already exists on every slide (for example, because the
template's slide master has a placeholder for it), iterate the
placeholders by type or name instead of creating a new textbox::

    from pptx import Presentation

    PROJECT_NAME = "Apollo — Q3 2026 review"

    prs = Presentation("template.pptx")

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.name == "ProjectHeader":      # authored on the master
                shape.text_frame.text = PROJECT_NAME

    prs.save("out.pptx")

And if the string needs to update *later* — the classic "the customer
renamed the project at the last minute" scenario — re-run the script
against the already-generated deck. python-pptx's
:meth:`.TextFrame.replace_text` and :meth:`.Picture.replace_image` are
the token-based variant of the same pattern (see
:ref:`template-style-replacement-text-and-pictures`): leave
``{{ project_name }}`` markers in the template and call
``text_frame.replace_text("{{ project_name }}", new_name)`` on each
text frame at generation time.


Why not author a ``<a:fld>`` of a custom type?
----------------------------------------------

Nothing *stops* you from writing an ``<a:fld>`` with a made-up ``@type``
attribute — the XML is syntactically legal. But the ``@type`` values
PowerPoint recognises are hard-coded in its binary; any other value is
treated as an opaque literal and the ``<a:t>`` child is displayed
verbatim. There is no consumer that will resolve a custom ``@type`` to a
string taken from another slide, so authoring one buys you nothing over
writing a plain run of text. Keeping the generation script as the
source of truth is simpler, more portable, and compatible with every
PowerPoint consumer (PowerPoint itself, Keynote, Google Slides,
LibreOffice Impress, and every viewer downstream).


See also
--------

* :doc:`slides` — the *Header, footer, slide number, and date
  placeholders* section — the one place OOXML *does* provide
  auto-refreshing fields.
* :doc:`ole-objects` — for live-linked Excel content, the nearest native
  analogue to cross-document text linking.
* :ref:`template-style-replacement-text-and-pictures` — token-based text
  replacement across a pre-authored template.
