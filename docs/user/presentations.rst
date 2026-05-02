
Working with Presentations
==========================

|pp| allows you to create new presentations as well as make changes to
existing ones. Actually, it only lets you make changes to existing
presentations; it's just that if you start with a presentation that doesn't
have any slides, it feels at first like you're creating one from scratch.

However, a lot of how a presentation looks is determined by the parts that are
left when you delete all the slides, specifically the theme, the slide master,
and the slide layouts that derive from the master. Let's walk through it a step
at a time using examples, starting with the two things you can do with
a presentation, open it and save it.


Opening a presentation
----------------------

The simplest way to get started is to open a new presentation without
specifying a file to open::

    from pptx import Presentation

    prs = Presentation()
    prs.save('test.pptx')

This creates a new presentation from the built-in default template and saves it
unchanged to a file named 'test.pptx'. A couple things to note:

* The so-called "default template" is actually just a PowerPoint file that
  doesn't have any slides in it, stored with the installed |pp| package. It's
  the same as what you would get if you created a new presentation from a fresh
  PowerPoint install, a 4x3 aspect ratio presentation based on the "White"
  template. Well, except it won't contain any slides. PowerPoint always adds
  a blank first slide by default.

* You don't need to do anything to it before you save it. If you want to see
  exactly what that template contains, just look in the 'test.pptx' file this
  creates.

* We've called it a *template*, but in fact it's just a regular PowerPoint file
  with all the slides removed. Actual PowerPoint template files (.potx files)
  are something a bit different. More on those later maybe, but you won't need
  them to work with |pp|.


Choosing a 16:9 (widescreen) or 4:3 default template
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The built-in default template is the classic 4:3 ``screen4x3`` layout (10 in
× 7.5 in, 9144000 × 6858000 EMU). If you'd rather start from a widescreen
16:9 deck (13.333 in × 7.5 in, 12192000 × 6858000 EMU), pass the
``pptx_format`` keyword argument::

    from pptx import Presentation

    prs = Presentation(pptx_format="16x9")          # widescreen
    prs = Presentation(pptx_format="widescreen")    # same as above
    prs = Presentation(pptx_format="4x3")           # explicit 4:3
    prs = Presentation(pptx_format="standard")      # same as "4x3"

The accepted values are case-insensitive and only recognised when no path
(or file-like object) is supplied — the argument controls which built-in
template ships into the new deck. Passing ``pptx_format`` together with
a filename raises :class:`ValueError` because the slide size is determined
by the file you're opening in that case.

To override the slide size of an already-loaded deck (either the default
template or a user-provided file) set
:attr:`~pptx.presentation.Presentation.slide_width` and
:attr:`~pptx.presentation.Presentation.slide_height` directly::

    from pptx.util import Emu

    prs = Presentation()                   # default 4:3
    prs.slide_width = Emu(12192000)        # switch to 16:9
    prs.slide_height = Emu(6858000)

Note that only the top-level slide size changes in that case; the
placeholders defined on the built-in slide master still reflect a 4:3
layout, so visible placeholders will be left-anchored within the wider
slide. For a clean widescreen starting point, prefer
``Presentation(pptx_format="16x9")``.


REALLY opening a presentation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Okay, so if you want any control at all to speak of over the final
presentation, or if you want to change an existing presentation, you need to
open one with a filename::

    prs = Presentation('existing-prs-file.pptx')
    prs.save('new-file-name.pptx')

Things to note:

* You can open any PowerPoint 2007 or later file this way (.ppt files from
  PowerPoint 2003 and earlier won't work). While you might not be able to
  manipulate all the contents yet, whatever is already in there will load and
  save just fine. The feature set is still being built out, so you can't add or
  change things like Notes Pages yet, but if the presentation has them |pp| is
  polite enough to leave them alone and smart enough to save them without
  actually understanding what they are.

* In addition to the regular presentation format (``.pptx``), |pp| also accepts
  macro-enabled presentations (``.pptm``), PowerPoint template files
  (``.potx``), and slideshow files (``.ppsx``). These are all recognized as
  PresentationML "main document" content types and open the same way -- just
  pass the path (or a file-like object) to ``Presentation()``. Note that |pp|
  always saves as a regular ``.pptx`` package; if you want to preserve the
  original template/slideshow content type you will need to rewrite the
  ``[Content_Types].xml`` entry yourself after saving.

* If you use the same filename to open and save the file, |pp| will obediently
  overwrite the original file without a peep. You'll want to make sure that's
  what you intend.


Opening a 'file-like' presentation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

|pp| can open a presentation from a so-called *file-like* object. It can also
save to a file-like object. This can be handy when you want to get the source
or target presentation over a network connection or from a database and don't
want to (or aren't allowed to) fuss with interacting with the file system. In
practice this means you can pass an open file or StringIO/BytesIO stream object
to open or save a presentation like so::

    f = open('foobar.pptx')
    prs = Presentation(f)
    f.close()

    # or

    with open('foobar.pptx') as f:
        source_stream = StringIO(f.read())
    prs = Presentation(source_stream)
    source_stream.close()
    ...
    target_stream = StringIO()
    prs.save(target_stream)


Reproducible saves (source-control friendly)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

By default, each saved ``.pptx`` stamps every zip member with the current
wall-clock time. That makes two otherwise-identical saves differ byte-for-byte
and show up as changed in git even though the underlying document is the same.

Pass ``zip_date_time`` to :meth:`~pptx.presentation.Presentation.save` to fix
every member's last-modified timestamp to a constant value — the resulting
``.pptx`` is then byte-identical across repeated saves of identical content::

    import datetime
    prs.save('deck.pptx', zip_date_time=datetime.datetime(2020, 1, 1))

    # or as a 6-tuple (year, month, day, hour, minute, second)
    prs.save('deck.pptx', zip_date_time=(2020, 1, 1, 0, 0, 0))

The earliest representable Zip date is 1980-01-01. The value may be a
|datetime| (tzinfo is ignored) or a 6-tuple of ints. When ``zip_date_time`` is
omitted, the legacy current-time behavior is preserved.


Opening and saving a password-protected presentation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

|pp| can read and write password-protected (encrypted) ``.pptx`` files that use
the ECMA-376 Agile Encryption scheme PowerPoint writes when a user chooses
"Encrypt with Password" in the desktop application. Pass the password via the
``password`` keyword argument::

    prs = Presentation('secret.pptx', password='hunter2')
    # ... edit as normal ...
    prs.save('new-secret.pptx', password='hunter2')

This feature requires the optional ``msoffcrypto-tool`` package::

    pip install msoffcrypto-tool

If a password-protected file is opened without a password, or with the wrong
password, a :class:`pptx.exc.EncryptedPackageError` is raised. Omitting the
``password`` argument when saving produces a plain (unencrypted) ``.pptx`` as
before.

The full keyword signature of :meth:`~pptx.presentation.Presentation.save` is
therefore ``save(file, *, zip_date_time=None, password=None)``. The two
keywords are orthogonal — you can use them together::

    import datetime
    prs.save(
        'deck.pptx',
        zip_date_time=datetime.datetime(2020, 1, 1),
        password='hunter2',
    )

When both are supplied, python-pptx first builds the plaintext zip with the
fixed timestamp stamped on every member and then wraps it in the ECMA-376
Agile Encryption CFBF container before writing. The resulting encrypted
``.pptx`` is *not* byte-reproducible (the encryption container itself carries
randomly-derived key-material and IV bytes), but decrypting it yields a
byte-reproducible inner zip.


Okay, so you've got a presentation open and are pretty sure you can save it
somewhere later. Next step is to get a slide in there ...


Extended document properties are synced on save
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Every ``.pptx`` package carries an *extended properties* XML part at
``docProps/app.xml`` that records (among other things) the total slide count
in a ``<Slides>`` element. Gmail's attachment-preview feature and a number of
third-party thumbnail / HTML renderers look at that value and will silently
fail to render a preview when it disagrees with the actual number of slides.

|pp| takes care of this for you: each time you call :meth:`Presentation.save`,
the ``<Slides>`` count is recomputed from the presentation's live slide list
and written back before serialization. If the source package does not have an
extended-properties part (uncommon; the built-in template ships with one),
a minimal one is added automatically.

You do not need to do anything for this behaviour — there is no API to opt
out, and the other application-level fields (``<Application>``, ``<Company>``,
``<TitlesOfParts>``, …) are not modified.
