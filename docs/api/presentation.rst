
Presentations
=============

A presentation is opened using the :func:`Presentation` function, provided
directly by the :mod:`pptx` package::

    from pptx import Presentation


This function returns a :class:`.Presentation` object which is the root of
a graph containing the components that constitute a presentation, e.g.
slides, shapes, etc. All existing presentation components are referenced by
traversing the graph and new objects are added to the graph by calling
a method on that object's container. Consequently, |pp| objects are generally
not constructed directly.

Example::

   # load a presentation
   prs = Presentation(path_to_pptx_file)

   # get reference to first shape in first slide
   sp = prs.slides[0].shapes[0]

   # add a picture shape to slide
   pic = sld.shapes.add_picture(path, x, y, cx, cy)


|Presentation| function
-----------------------

This function is the only reference that must be imported to work with
presentation files. Typical use interacts with many other classes, but there
is no need to construct them as they are accessed using a property or method
of their containing object.

.. autofunction:: pptx.Presentation


|Presentation| objects
-----------------------

.. autoclass:: pptx.presentation.Presentation()
   :members:
   :member-order: bysource
   :exclude-members: part
   :undoc-members:


Embedding custom fonts
----------------------

The :meth:`Presentation.embed_font` method embeds a TrueType or OpenType font
file in the presentation so that text using that typeface renders correctly
on machines where the font is not installed. PowerPoint supports up to four
style "slots" per typeface — *regular*, *bold*, *italic*, and *boldItalic* —
and :meth:`embed_font` can be called repeatedly to populate additional
slots for the same typeface.

Example::

    from pptx import Presentation

    prs = Presentation()
    prs.embed_font("fonts/Pacifico-Regular.ttf", "Pacifico")
    prs.embed_font("fonts/Pacifico-Bold.ttf", "Pacifico", style="bold")
    print(prs.embedded_fonts)   # -> ('Pacifico',)
    prs.save("branded.pptx")

``embed_font()`` accepts either a filesystem path or any file-like object
opened for binary reading, so an in-memory ``BytesIO`` also works. Licensing
for font embedding is the caller's responsibility; see your font's
end-user license agreement.


|CoreProperties| objects
-------------------------

Each |Presentation| object has a |CoreProperties| object accessed via its
:attr:`core_properties` attribute that provides read/write access to the
so-called *core properties* for the document. The core properties are author,
category, comments, content_status, created, identifier, keywords, language,
last_modified_by, last_printed, modified, revision, subject, title, and
version.

Each property is one of three types, |str|, |datetime|, or |int|. String
properties are limited in length to 255 characters and return an empty string
('') if not set. Date properties are assigned and returned as |datetime|
objects without timezone, i.e. in UTC. Any timezone conversions are the
responsibility of the client. Date properties return |None| if not set.

|pp| does not automatically set any of the document core properties other than
to add a core properties part to a presentation that doesn't have one (very
uncommon). If |pp| adds a core properties part, it contains default values for
the title, last_modified_by, revision, and modified properties. Client code
should change properties like revision and last_modified_by explicitly if that
behavior is desired.

.. class:: pptx.opc.coreprops.CoreProperties

   .. attribute:: author

      *string* -- An entity primarily responsible for making the content of the
      resource.

   .. attribute:: category

      *string* -- A categorization of the content of this package. Example
      values might include: Resume, Letter, Financial Forecast, Proposal,
      or Technical Presentation.

   .. attribute:: comments

      *string* -- An account of the content of the resource.

   .. attribute:: content_status

      *string* -- completion status of the document, e.g. 'draft'

   .. attribute:: created

      *datetime* -- time of intial creation of the document

   .. attribute:: identifier

      *string* -- An unambiguous reference to the resource within a given
      context, e.g. ISBN.

   .. attribute:: keywords

      *string* -- descriptive words or short phrases likely to be used as
      search terms for this document

   .. attribute:: language

      *string* -- language the document is written in

   .. attribute:: last_modified_by

      *string* -- name or other identifier (such as email address) of person
      who last modified the document

   .. attribute:: last_printed

      *datetime* -- time the document was last printed

   .. attribute:: modified

      *datetime* -- time the document was last modified

   .. attribute:: revision

      *int* -- number of this revision, incremented by the PowerPoint® client
      once each time the document is saved. Note however that the revision
      number is not automatically incremented by |pp|.

   .. attribute:: subject

      *string* -- The topic of the content of the resource.

   .. attribute:: title

      *string* -- The name given to the resource.

   .. attribute:: version

      *string* -- free-form version string


Extended (application) properties
---------------------------------

The *extended-properties* part (``/docProps/app.xml``) contains
application-level metadata about the presentation, including a ``<Slides>``
element that records the total slide count. Several downstream consumers,
notably Gmail's attachment-preview feature and various third-party
thumbnail / HTML renderers, refuse to render the presentation when
``<Slides>`` disagrees with the actual number of slides.

|pp| automatically recomputes and writes the ``<Slides>`` count on each
:meth:`Presentation.save` call, using the length of the internal
``sldIdLst``. If the package has no extended-properties part (uncommon;
the default template includes one), a minimal part with only the
``<Slides>`` element is added.

This behavior is intentionally transparent — there is no public API to
opt out, and the other application-level fields (``<Application>``,
``<Company>``, ``<TitlesOfParts>``, …) are left untouched if present and
are not populated if absent.
