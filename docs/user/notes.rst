
Working with Notes Slides
=========================

A slide can have notes associated with it. These are perhaps most commonly
encountered in the notes pane, below the slide in PowerPoint "Normal" view
where it may say "Click to add notes".

The notes added here appear each time that slide is present in the main pane.
They also appear in *Presenter View* and in *Notes Page* view, both available
from the menu.

Notes can contain rich text, commonly bullets, bold, varying font sizes and
colors, etc. The Notes Page view has somewhat more powerful tools for editing
the notes text than the note pane in Normal view.

In the API and the underlying XML, the object that contains the text is known
as a *Notes Slide*. This is because internally, a notes slide is actually
a specialized instance of a slide. It contains shapes, many of which are
placeholders, and allows inserting of new shapes such as pictures (a logo
perhaps) auto shapes, tables, and charts. Consequently, working with a notes
slide is very much like working with a regular slide.

Each slide can have zero or one notes slide. A notes slide is created the
first time it is used, generally perhaps by adding notes text to a slide.
Once created, it stays, even if all the text is deleted.

The Notes Master
----------------

A new notes slide is created using the *Notes Master* as a template.
A presentation has no notes master when newly created in PowerPoint. One is
created according to a PowerPoint-internal preset default the first time it
is needed, which is generally when the first notes slide is created. It's
possible one can also be created by entering the Notes Master view and almost
certainly is created by editing the master found there (haven't tried it
though). A presentation can have at most one notes master.

The notes master governs the look and feel of notes pages, which can be
viewed on-screen but are really designed for printing out. So if you want
your notes page print-outs to look different from the default, you can make
a lot of customizations by editing the notes master. You access the notes
master editor using View > Master > Notes Master on the menu (on my version
at least). Notes slides created using |pp| will have the look and feel of the
notes master in the presentation file you opened to create the presentation.

On creation, certain placeholders (slide image, notes, slide number) are
copied from the notes master onto the new notes slide (if they have not been
removed from the master). These "cloned" placeholders inherit position, size,
and formatting from their corresponding notes master placeholder. If the
position, size, or formatting of a notes slide placeholder is changed, the
changed property is no long inherited (unchanged properties, however,
continue to be inherited).

Notes Slide basics
------------------

Enough talk, let's show some code. Let's say you have a slide you're working
with and you want to see if it has a notes slide yet::

    >>> slide.has_notes_slide
    False

Ok, not yet. Good. Let's add some notes::

    >>> notes_slide = slide.notes_slide
    >>> text_frame = notes_slide.notes_text_frame
    >>> text_frame.text = 'foobar'

Alright, simple enough. Let's look at what happened here:

* ``slide.notes_slide`` gave us the notes slide. In this case, it first
  created that notes slide based on the notes master. If there was no notes
  master, it created that too. So a lot of things can happen behind the
  scenes with this call the first time you call it, but if we called it again
  it would just give us back the reference to the same notes slide, which it
  caches, once retrieved.

* ``notes_slide.notes_text_frame`` gave us the |TextFrame| object that
  contains the actual notes. The reason it's not just
  ``notes_slide.text_frame`` is that there are potentially more than one. What
  this is doing behind the scenes is finding the placeholder shape that
  contains the notes (as opposed to the slide image, header, slide number,
  etc.) and giving us *that* particular text frame.

* A text frame in a notes slide works the same as one in a regular slide.
  More precisely, a text frame on a shape in a notes slide works the same as
  in any other shape. We used the ``.text`` property to quickly pop some text
  in there.

Using the text frame, you can add an arbitrary amount of text, formatted
however you want.


Notes Slide Placeholders
------------------------

What we haven't explicitly seen so far is the shapes on a slide master. It's
easy to get started with that::

    >>> notes_placeholder = notes_slide.notes_placeholder

This notes placeholder is just like a body placeholder we saw a couple
sections back. You can change its position, size, and many other attributes,
as well as get at its text via its text frame.

You can also access the other placeholders::

    >>> for placeholder in notes_slide.placeholders:
    ...   print placeholder.placeholder_format.type
    ...
    SLIDE_IMAGE (101)
    BODY (2)
    SLIDE_NUMBER (13)

and also the shapes (a superset of the placeholders)::

    >>> for shape in notes_slide.shapes:
    ...   print shape
    ...
    <pptx.shapes.placeholder.NotesSlidePlaceholder object at 0x11091e890>
    <pptx.shapes.placeholder.NotesSlidePlaceholder object at 0x11091e750>
    <pptx.shapes.placeholder.NotesSlidePlaceholder object at 0x11091e990>

In the common case, the notes slide contains only placeholders. However, if
you added an image, for example, to the notes slide, that would show up as
well. Note that if you added that image to the notes master, perhaps a logo,
it would appear on the notes slide "visually", but would not appear as
a shape in the notes slide shape collection. Rather, it is visually
"inherited" from the notes master.
