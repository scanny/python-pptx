
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


Choosing a slide size for the default template
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The built-in default template is the classic 4:3 ``screen4x3`` layout (10 in
× 7.5 in, 9144000 × 6858000 EMU). To start from a different slide size, pass
the ``pptx_format`` keyword argument — either a named preset or a raw
``(cx, cy)`` tuple in English Metric Units::

    from pptx import Presentation

    prs = Presentation(pptx_format="16x9")          # widescreen (13.333 x 7.5 in)
    prs = Presentation(pptx_format="widescreen")    # same as above
    prs = Presentation(pptx_format="4x3")           # explicit 4:3 (10 x 7.5 in)
    prs = Presentation(pptx_format="standard")      # same as "4x3"
    prs = Presentation(pptx_format="letter")        # US Letter (10 x 7.5 in, tagged ``letter``)
    prs = Presentation(pptx_format="a4")            # A4 landscape (297 x 210 mm)

    # arbitrary size via (cx, cy) in EMU
    prs = Presentation(pptx_format=(10058400, 7772400))   # 11 x 8.5 in

Preset names are case-insensitive. The ``letter`` and ``a4`` presets set the
ECMA-376 ``p:sldSz/@type`` attribute to ``letter`` / ``A4`` so PowerPoint's
Page Setup dialog reports the slide as letter- or A4-paper formatted. A
``(cx, cy)`` tuple sets ``p:sldSz/@type`` to ``custom``.

``pptx_format`` is only recognised when no path (or file-like object) is
supplied — it controls which built-in template ships into the new deck.
Passing ``pptx_format`` together with a filename raises
:class:`ValueError` because the slide size is determined by the file you're
opening in that case.

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


Using an existing presentation as a blank template
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

A common workflow is to start a new deck from a branded, corporate-style
``.pptx`` file — keep the slide master, slide layouts, theme, embedded
fonts, and default table styles, but throw away the slides that happened
to be in the file when it was saved. The
:meth:`~pptx.presentation.Presentation.strip_slides` method does exactly
that::

    from pptx import Presentation

    blank = Presentation("branded-template.pptx").strip_slides()

    # -- blank is the same Presentation instance, now with no slides but
    # -- the masters / layouts / theme / fonts from branded-template.pptx
    # -- intact. Add slides like you would to a fresh deck. --
    title_slide = blank.slides.add_slide(blank.slide_layouts[0])
    title_slide.shapes.title.text = "Quarterly review"

    blank.save("q3-review.pptx")

``strip_slides()`` returns ``self`` so it composes with the
:func:`Presentation` factory call. The operation also clears any
presentation sections the source file carried, since a section is just a
named group of slide-ids and an empty section list with stale
slide-id references would confuse PowerPoint. Slide parts, uniquely-
referenced image/media/chart parts, and the section scaffolding all
become unreachable from the package root and are omitted on the next
save. Parts shared with the surviving slide master or layouts are
retained, so theme assets and template-level resources are preserved.

Calling ``strip_slides()`` on a deck that already has no slides is a
no-op. See issue #310 for the original request.


Okay, so you've got a presentation open and are pretty sure you can save it
somewhere later. Next step is to get a slide in there ...


Organizing slides into sections
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Starting with PowerPoint 2010, a presentation can be partitioned into
named *sections* that group consecutive slides together in the
Navigation pane. |pp| exposes these through the
:attr:`~pptx.presentation.Presentation.sections` collection::

    from pptx import Presentation

    prs = Presentation("deck.pptx")

    # -- add a section containing the first two slides --
    intro = prs.sections.add_section(
        "Intro", slides=[prs.slides[0], prs.slides[1]]
    )

    # -- read-only id, read/write name --
    print(intro.id)        # -> '{EA541957-60A5-47A7-8DF8-F51D8F183008}'
    intro.name = "Overview"

    # -- assign / revoke section membership for individual slides --
    intro.add_slide(prs.slides[2])
    intro.remove_slide(prs.slides[0])

    # -- look up by name or id --
    prs.sections.get_by_name("Overview")
    prs.sections.get_by_id(intro.id)

    # -- tear it down --
    prs.sections.remove(intro)

A few things to note:

* Each section carries a stable GUID ``id``. |pp| generates one
  automatically; pass ``id="{...}"`` to :meth:`Sections.add_section` if
  you need a specific value (e.g. preserving ids from another source).
* A slide can belong to at most one section in PowerPoint. |pp| does
  **not** automatically move a slide out of another section when you
  call :meth:`Section.add_slide` on a new section; if you need that
  mutual-exclusion behavior, remove the slide from the prior section
  first.
* Removing the last section also prunes the empty ``p:extLst``
  scaffolding so a round-tripped deck does not retain an empty
  extension block.


Merging two presentations
~~~~~~~~~~~~~~~~~~~~~~~~~

:meth:`Presentation.merge` appends every slide of one presentation to another
in a single call. The copied slides are full-fidelity duplicates — shape tree,
image and media parts, charts (each with its own embedded workbook), embedded
OLE objects, and external hyperlinks are all materialised in the target
presentation's package. The source presentation is not modified::

    from pptx import Presentation

    target = Presentation("report-shell.pptx")
    source = Presentation("quarterly-results.pptx")

    # -- append every slide of `source` onto `target` --
    appended = target.merge(source)

    print(len(appended), "slides merged")
    target.save("combined.pptx")

Things to note:

* Each cloned slide binds to the layout at the **same index** in the target
  master's layout list as its source's layout occupied in the source master.
  When the target master has fewer layouts, the last layout is used as a
  fallback. Callers who need strict layout mapping should reassign
  ``slide.slide_layout`` after the merge or use
  :meth:`Slides.add_slide_from_external` for per-slide control.
* Image and media parts are content-deduplicated against the target package;
  an image already present is reused rather than duplicated.
* Each chart carries across with its *own* embedded workbook so PowerPoint's
  "Edit Data" dialog continues to work on both the original and the merged copy.
* Notes slides are **dropped** on the copies: a notes slide in OOXML stores a
  back-reference to its owning slide and so cannot be shared. Add fresh notes
  on the appended slides as needed.

:meth:`Presentation.merge` raises ``TypeError`` when passed a non-
|Presentation| argument and ``ValueError`` if called with ``self`` as the
argument (to prevent inadvertently doubling a deck's slide count).


Applying a POTX/PPTX template to existing content
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

A recurring request (issue #1095) is to apply a corporate ``.potx`` — or
any branded ``.pptx`` — to a deck of existing content slides, the way
PowerPoint's "Design → Reuse Slides / Apply Template" flow does. |pp|
does not expose a single "apply template" call, but the three pieces
needed to express the workflow are each first-class:

* opening a PowerPoint template (``.potx``) or slideshow (``.ppsx``)
  package with :func:`Presentation` (see `REALLY opening a presentation`_),
* :meth:`~pptx.presentation.Presentation.strip_slides` to keep only the
  template's masters / layouts / theme,
* :meth:`~pptx.presentation.Presentation.merge` to graft each content
  slide into the stripped template deck.

Composed, the recipe is::

    from pptx import Presentation

    # -- 1. Open the branded template and keep only its masters / layouts. --
    target = Presentation("brand.potx").strip_slides()

    # -- 2. Open the source content deck. --
    content = Presentation("quarterly-results.pptx")

    # -- 3. Append every content slide into the template deck. --
    target.merge(content)

    # -- 4. Save as a regular .pptx (see note on content type below). --
    target.save("quarterly-results-rebranded.pptx")

What this preserves and what it does not:

* **Preserved from the template:** slide masters, slide layouts, theme,
  embedded fonts, default table styles, and any other template-level
  resources in ``brand.potx``.
* **Preserved from the content:** each slide's shape tree, image/media
  parts, charts (each retaining its own embedded workbook), embedded
  OLE objects, and external hyperlinks — per :meth:`Presentation.merge`.
* **Layout binding:** each cloned slide binds to the layout at the *same
  index* in the target master as its source's layout occupied in the
  source master. If the template has fewer layouts the last layout is
  used as a fallback. When you need strict layout mapping — e.g.
  re-mapping a "Title and Content" slide from the source onto a named
  "Content" layout in the template — iterate the source slides and use
  :meth:`Slides.add_slide_from_external` with an explicit target layout
  instead of :meth:`merge`.
* **Notes slides are dropped** on merged slides (a notes slide back-
  references its owning slide and so cannot be shared) — re-attach
  speaker notes after the merge if you need them.
* **Output file extension / content type:** :meth:`Presentation.save`
  always writes a regular ``.pptx`` package. If you opened a ``.potx``
  for the template and want the result to remain a template, you will
  need to rewrite the ``[Content_Types].xml`` main-document entry
  yourself after saving — see the opening-a-presentation note above.

Calling ``strip_slides()`` before ``merge()`` is important. Skipping it
would leave the slides that happened to ship inside the ``.potx`` (most
corporate templates include a title / agenda sample) sitting ahead of
the merged content in the output deck.


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


Custom document properties (DOCPROPERTY fields)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

``Presentation.custom_properties`` exposes the user-defined "Custom"
document properties stored in ``/docProps/custom.xml`` — the properties
PowerPoint surfaces under *File → Info → Properties → Advanced → Custom*
and that ``{ DOCPROPERTY MyProp }`` field codes resolve against. The
facade is dict-like, lazily materialises the underlying part (reading
from an unlabelled deck never forces the part into the package), and
preserves python types across save/reopen::

    from pptx import Presentation
    import datetime as dt

    prs = Presentation()

    prs.custom_properties["Department"] = "Engineering"
    prs.custom_properties["Revision"] = 7
    prs.custom_properties["Score"] = 98.6
    prs.custom_properties["Approved"] = True
    prs.custom_properties["Deadline"] = dt.datetime(2025, 3, 15, 12, 0, 0)

    prs.save("deck.pptx")

Supported value types are ``str`` (serialised as ``<vt:lpwstr>``),
``int`` (``<vt:i4>``), ``float`` (``<vt:r8>``), ``bool``
(``<vt:bool>``), and :class:`datetime.datetime` (``<vt:filetime>``).
The facade also implements the usual ``dict`` operations — ``in``,
``len()``, ``keys()``, ``items()``, ``update()``, ``pop()``,
``setdefault()``, ``clear()``, and ``del``. See issue #259.


Microsoft sensitivity labels (MIP)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Microsoft Information Protection (MIP) / Azure Information Protection
(AIP) sensitivity labels — the "Confidential", "Internal", "Public" tags
PowerPoint's *Sensitivity* menu applies — are not a first-class OOXML
feature. They ride on top of the same ``/docProps/custom.xml`` part
covered in the section above, as a conventional bundle of custom
properties whose names follow the pattern::

    MSIP_Label_<GUID>_<Field>

For each label applied to a presentation, PowerPoint emits six fields:

============================ =================================================================
Field suffix                 Value
============================ =================================================================
``_Enabled``                 ``"true"`` / ``"false"`` (stored as a string, not a bool)
``_SetDate``                 ISO-8601 UTC timestamp, e.g. ``"2025-05-02T09:15:00Z"``
``_Method``                  ``"Standard"`` or ``"Privileged"``
``_Name``                    Human-readable label name (``"Confidential"`` etc.)
``_SiteId``                  Tenant GUID, in the same brace-enclosed form as ``<GUID>``
``_ContentBits``             Integer bit-field flagging where the label is visualised
                             (0 = none, 1 = header, 2 = footer, 4 = watermark)
============================ =================================================================

The ``<GUID>`` segment is the label's unique identifier, assigned by the
administrator in the Microsoft Purview / AIP console, and is the same
value for every field in the bundle. Write the bundle directly through
:attr:`~pptx.presentation.Presentation.custom_properties`::

    from pptx import Presentation

    prs = Presentation("unlabelled.pptx")

    guid = "{defa4170-0d19-0005-0004-bc88714345d2}"   # "Confidential" in this tenant
    site_id = "{72f988bf-86f1-41af-91ab-2d7cd011db47}"

    prefix = "MSIP_Label_%s" % guid
    prs.custom_properties["%s_Enabled" % prefix] = "true"
    prs.custom_properties["%s_SetDate" % prefix] = "2025-05-02T09:15:00Z"
    prs.custom_properties["%s_Method" % prefix] = "Standard"
    prs.custom_properties["%s_Name" % prefix] = "Confidential"
    prs.custom_properties["%s_SiteId" % prefix] = site_id
    prs.custom_properties["%s_ContentBits" % prefix] = 0

    prs.save("labelled.pptx")

Things to note:

* ``_Enabled`` is a **string**, not a Python ``bool`` — that matches
  what the MIP / AIP clients and downstream DLP scanners parse. Setting
  it to ``True`` would serialise as ``<vt:bool>`` instead of the
  expected ``<vt:lpwstr>true</vt:lpwstr>`` and break recognition.
* ``_SetDate`` is likewise a **string** (ISO-8601 text), not a
  :class:`datetime.datetime`. PowerPoint writes it as
  ``<vt:lpwstr>`` and DLP scanners read it that way.
* ``_ContentBits`` **is** an integer, and serialises as ``<vt:i4>``.
  Set it to ``0`` if the label is metadata-only (no visible
  header/footer/watermark).
* The GUID is persistent across tenants for standard labels
  (e.g. ``defa4170-0d19-0005-0004-bc88714345d2`` is Microsoft's
  canonical "Confidential" id) but for custom labels you must obtain it
  from the tenant administrator. There is no library-side GUID
  allocation — if you invent one, the label will not round-trip through
  a MIP-aware application.
* Multiple labels can coexist (distinct GUIDs), which is sometimes seen
  during migration between label taxonomies. Consumers typically act on
  the most recent ``_SetDate`` among enabled labels.

Reading labels back is symmetrical — enumerate the custom-properties
keys and pick off the ``MSIP_Label_<GUID>_*`` family::

    def applied_label_guids(prs):
        """Return the set of MIP-label GUIDs carried by `prs`."""
        seen = set()
        for key in prs.custom_properties.keys():
            if key.startswith("MSIP_Label_"):
                tail = key[len("MSIP_Label_"):]
                guid, _, _field = tail.rpartition("_")
                if guid:
                    seen.add(guid)
        return seen

Clearing a label is equally direct — delete every key in its bundle::

    prefix = "MSIP_Label_%s_" % guid
    for key in [k for k in prs.custom_properties.keys() if k.startswith(prefix)]:
        del prs.custom_properties[key]

python-pptx exposes the underlying :attr:`custom_properties` mapping
rather than a dedicated sensitivity-label class so the recipe stays
forward-compatible with any MIP field the Microsoft tooling may add in
future. See issue #882.
