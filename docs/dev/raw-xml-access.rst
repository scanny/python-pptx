Raw XML access (beyond-API escape hatches)
==========================================

|pp| aims to be a *general-purpose* library, but PowerPoint's OOXML grammar is
enormous and a given release will never expose a Python wrapper for every
element or attribute the schema defines. When you need to manipulate XML that
the public API has not yet modelled — an obscure attribute, a Microsoft
extension namespace (``p14:``, ``a14:``, ``p15:``, ``cx:``, …), or a round-trip
tweak to support a corner case — |pp| lets you drop down to the underlying
``lxml`` tree without forking the library.

This page documents the supported escape hatches for direct XML access and
shows how to use the ``xmlchemy`` descriptor layer to do so in a typesafe way
whenever possible. Addresses GitHub issue #626.

.. contents::
   :local:
   :depth: 2


When to reach for raw XML
-------------------------

Prefer the public API whenever it covers your need. Raw-XML edits:

* bypass the type- and namespace-checking the ``xmlchemy`` descriptors enforce;
* can leave the package in a state PowerPoint refuses to open;
* must be re-reviewed every time you upgrade |pp|, because a future release may
  model the element you're touching and subtly change its behaviour.

Still, raw XML access is a first-class supported feature. It is the intended
escape hatch for:

* **Reading** attributes or child elements that |pp| does not surface yet.
* **Round-trip preservation** when you need to tweak something |pp| would
  otherwise leave alone.
* **Experimentation** before proposing a feature — reach into the XML, confirm
  the behaviour, then lift the pattern up into a proper API addition.

If you find yourself reaching for this escape hatch repeatedly for the same
element, that's a strong signal the library should model it. Please open an
issue (or a PR using the ``xmlchemy`` layer; see
:doc:`xmlchemy` and :doc:`development_practices`).


The ``element`` property
------------------------

Most proxy classes derive from :class:`pptx.shared.ElementProxy`, which exposes
the underlying ``lxml`` element via a public ``element`` property::

    shape_elm = shape.element           # a CT_Shape / CT_Picture / …
    slide_elm = slide.element           # a CT_Slide (the `p:sld` root)
    prs_elm   = prs.element             # a CT_Presentation
    section_elm = section.element       # a CT_Section

``BaseShape`` (the root of the shape class hierarchy) also implements
``.element`` so every shape subclass — ``Shape``, ``Picture``, ``GraphicFrame``,
``Connector``, ``GroupShape``, and the placeholder variants — participates in
the same pattern::

    shape_elm = shape.element
    # -> ShapeElement instance (lxml Element subclass with xmlchemy behaviour)

The returned object is a live reference. Mutating it in place mutates the
presentation; you do not need to write anything back. When you eventually call
``prs.save(...)`` the current state of the tree is serialized.

The ``_element`` attribute used internally by proxy classes is the same object
and can also be used, but ``.element`` is the preferred public accessor.


Chart: ``_chartSpace``
~~~~~~~~~~~~~~~~~~~~~~

A :class:`~pptx.chart.chart.Chart` wraps the ``c:chartSpace`` root of its
chart part. Because ``Chart`` derives from :class:`PartElementProxy`, the
``c:chartSpace`` is available as ``chart.element``. Internally it is also
bound to ``chart._chartSpace``::

    chartSpace = chart.element           # preferred
    chartSpace = chart._chartSpace       # same object (internal name)

From there the usual ``xmlchemy`` accessors reach the plot area, series,
axes, and the Microsoft extension subtree (``cx:`` elements appear inside
``mc:AlternateContent`` / ``mc:Choice`` wrappers).


Text: ``_txBody``, ``_p``, ``_r``
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Text-frame, paragraph, and run objects also keep references to their XML
elements. The shapes hierarchy carries a ``text_frame`` which itself exposes
its ``_txBody`` (``a:txBody``); paragraphs and runs carry their ``a:p`` and
``a:r`` elements respectively. These are internal names; callers who want to
manipulate them can use ``run.element``, ``paragraph.element``, and
``text_frame.element`` when available, or fall back to the underscore-prefixed
attributes for elements without a public accessor.


Part-level: ``.part.element``
-----------------------------

Every proxy that is attached to a package part exposes ``.part``. The part's
own root element — for an ``XmlPart`` — is reachable as ``part.element``. This
gives you access to the document-level XML for any part the library loads or
creates, including parts for which the proxy object does not expose
everything you need::

    slide_part_elm = slide.part.element      # the `p:sld` root
    chart_part_elm = chart.part.element      # the `c:chartSpace` root

This is the hook you use when you need to walk relationships, find sibling
parts, or write XML the part-level proxy does not model.


XPath search: ``Slide.find_shapes_by_xpath``
--------------------------------------------

For many "access XML the API doesn't expose" tasks, the shortest path is an
XPath query. :meth:`Slide.find_shapes_by_xpath` (added in this fork; see
HISTORY — issue #224) evaluates an XPath expression against the slide's
``p:spTree`` using the standard OOXML namespace map and returns the matching
shape proxies::

    # Every shape whose @name is "Title 1"
    shapes = slide.find_shapes_by_xpath(
        './/p:sp[p:nvSpPr/p:cNvPr/@name="Title 1"]'
    )

    # Every non-group shape anywhere in the tree
    shapes = slide.find_shapes_by_xpath(".//p:sp")

The standard OOXML prefixes (``p``, ``a``, ``r``, ``mc``, ``p14``, …) are
pre-registered, so you don't have to declare a namespace map yourself. Matches
that are themselves shape elements (``p:sp``, ``p:pic``, ``p:cxnSp``,
``p:graphicFrame``, ``p:grpSp``) are returned as their |pp| proxy;
sub-element matches resolve to their nearest shape ancestor. Duplicates are
suppressed.

For searches that must return non-shape elements (or that need to run against
a part other than a slide), fall back to ``lxml``'s native
:meth:`~lxml.etree._Element.xpath` on the element returned from ``.element``
and supply the namespace map from :data:`pptx.oxml.ns.nsmap` explicitly.


Pairing with the ``xmlchemy`` descriptor layer
----------------------------------------------

Raw ``lxml`` access is the escape hatch of last resort. Before writing custom
XPath or ``etree`` calls, check whether the element you need is already
modelled by a ``CT_*`` class in ``pptx.oxml``. If it is, you get type-safe
getters, setters, and insertion helpers for free:

* :class:`ZeroOrOne`, :class:`OneAndOnlyOne`, :class:`ZeroOrMore` — child
  element descriptors with correct namespace handling and insertion ordering.
* :class:`RequiredAttribute`, :class:`OptionalAttribute` — type-converted
  attribute access driven by ``ST_*`` simple types.

For instance, reaching for an ``a:xfrm`` child under a ``p:sp`` is just
``shape.element.spPr.xfrm`` — no ``findall``, no namespace-prefix juggling.

If the element is *not* modelled and you plan to reach for it often, consider
adding a ``CT_*`` class rather than sprinkling raw XPath through your code.
The :doc:`xmlchemy` page walks through the process.


Stability expectations
----------------------

The public ``.element`` accessor on :class:`ElementProxy` and :class:`BaseShape`
is a supported API and will not be renamed or removed without a deprecation
cycle. The underscore-prefixed internal attributes (``_element``, ``_chartSpace``,
``_txBody``, ``_prs_part``, etc.) are **not** part of the public API and may be
renamed or restructured between minor releases. Prefer ``.element`` where it is
available.

The shape of the underlying XML tree follows OOXML + PowerPoint's real output,
neither of which is strictly stable — but |pp|'s commitment to round-trip
fidelity means the XML you observe today is the XML |pp| will continue to
emit. Raw-XML edits that respect that shape remain valid across releases.
