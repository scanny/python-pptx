
Cross-part relationship cloner
==============================

Foundation item *F1* — a reusable helper that clones an OOXML element plus every
Part relationship it pulls with it (images, media, charts, embedded workbooks,
hyperlinks, OLE, ...). Exposed as
:class:`pptx.opc.package.PartRelationshipCloner` and underpins feature work on
slide duplication (`#132`, `#640`, `#835`), shape duplication (`#533`), chart
copy (`#338`, `#239`, `#833`), slide reorder/delete (`#68`, `#620`, `#879`),
picture-replace (`#116`), and combo-charts (`#877`).


Problem
-------

Several high-demand features all need the same primitive operation:

* Move or duplicate an XML element (a ``p:sld``, a ``p:pic``, a ``c:chart``, ...)
  from one part to another — possibly in the same package, possibly across
  packages.
* Bring along every *file-backed* thing the element points at via OPC
  relationships: images, media blobs, chart workbooks, OLE objects, hyperlinks.
* Do this without breaking rIds already in use on either side.

Before F1, every feature that touched cross-part copying was reinventing the
same wheel (see scattered one-off cases in ``ImagePart.new`` reuse,
``ChartWorkbook.xlsx_part`` setter, ``_ImageParts.get_or_add_image_part``).
These ad-hoc solutions didn't compose, and rId rewriting in particular was
error-prone because OOXML references rIds through three distinct attributes
(``r:id``, ``r:embed``, ``r:link``) on many element types.

The design goal is a single, narrow helper that handles the mechanical
bookkeeping — so higher-level code ("duplicate this shape", "clone this
slide") reads as intent, not plumbing.


Relationship reference attributes
---------------------------------

OOXML references Parts from XML content through three attributes, all in the
``r`` namespace (``http://schemas.openxmlformats.org/officeDocument/2006/relationships``):

``r:id``
   Generic target reference. Used by ``p:sldId`` (slide-in-presentation),
   ``c:chart`` (chart on a slide), ``a:hlinkClick`` (hyperlink), ``p:oleObj``,
   ``p:sldLayoutId``, ``p:sldMasterId``, and many more.

``r:embed``
   Embedded resource reference. Used on ``a:blip`` to point at an embedded
   image part.

``r:link``
   Linked (not embedded) resource reference. Used on ``a:videoFile``,
   ``p:audioFile``, and on ``a:blip`` when an image is linked rather than
   embedded. The referenced relationship typically has ``TargetMode="External"``.

Schemas: see ``spec/ISO-IEC-29500-1/schemas/xsd/dml-main.xsd`` (``a:blip``,
``a:hlinkClick``), ``pml.xsd`` (``p:sldId``, ``p:oleObj``), and
``spec/ISO-IEC-29500-2/opc-xsd/opc-relationships.xsd`` for the
``CT_Relationship`` grammar.

The cloner treats these three attributes uniformly: any element carrying one
of them is a reference, and every distinct rId value maps to exactly one
cloned relationship on the target part.


Algorithm
---------

Given ``src_part``, ``tgt_part``, and ``src_element``:

1. **Deep-copy** ``src_element``. Everything that follows mutates the copy —
   the source is never touched.
2. **Walk** the copy and its descendants in document order, collecting every
   distinct rId value found in ``r:id``/``r:embed``/``r:link`` attributes.
3. For each distinct rId:

   a. Resolve the corresponding ``_Relationship`` on ``src_part``.
   b. If it's **external** (``TargetMode="External"``), call
      ``tgt_part.relate_to(target_ref, reltype, is_external=True)``; this
      creates a fresh external relationship on ``tgt_part`` or returns an
      existing one pointing at the same URI.
   c. If it's **internal**, resolve ``rel.target_part``, then:

      * If the target part is already in the target package (same-package
        clone: the common case for slide duplication within a presentation),
        reuse it.
      * Otherwise, derive a non-colliding partname in the target package from
        the source partname template (e.g. ``/ppt/media/image3.png`` →
        ``/ppt/media/image%d.png`` → ``/ppt/media/image12.png``), construct
        a new ``Part`` of the same runtime class with the same content-type
        and blob, and call ``tgt_part.relate_to(new_part, reltype)``.

   d. Record ``{old_rId: new_rId}``.
4. **Rewrite** every rId attribute in the copy using the mapping.
5. **Return** the rewritten copy.

Content-type registration is implicit: ``PackageWriter`` derives
``[Content_Types].xml`` from the parts emitted by ``iter_parts()``, and
``iter_parts()`` in turn walks ``iter_rels()``. Adding a relationship from
``tgt_part`` to a new part therefore adds the new part to the package, which
adds its content-type. No explicit registration step is needed.


Deliberate non-goals
--------------------

* **Deep-recursive cloning** of a target part's own relationships. A chart
  part references an embedded ``.xlsx``; a slide references its slide-layout,
  which references its slide-master; cloning these graphs is the job of
  content-aware helpers composed on top of ``PartRelationshipCloner``. For
  example, F5 (cross-part embedded-workbook handler) layers chart-workbook
  cloning on top of this class.

* **Deduplicating** images or media by SHA-1 on the target side. Same-package
  cloning (where ``src_target_part.package is tgt_part.package``) already
  reuses the existing part; cross-package cloning creates a fresh part even
  if a blob-identical part already exists on the target. Deduplication is the
  job of ``Package._image_parts.get_or_add_image_part`` (by SHA-1), which
  callers can invoke explicitly when that's what they want.

* **Rewriting rIds** that resolve to a relationship missing on ``src_part``.
  In this case the cloner leaves the attribute unchanged (logged in code as
  ``# pragma: no cover`` — defensive). A missing relationship is a bug in
  the caller, not something to paper over silently.


Part helpers
------------

Two small methods on ``Part`` support the cloner:

``Part._new_rId()``
   Return an rId string that isn't yet used on the part. Thin wrapper over
   ``_Relationships._next_rId``; useful to callers that need to know the rId
   *before* the relationship is added (rare — the typical pattern is to let
   ``relate_to`` allocate the rId).

``Part._next_partname(tmpl)``
   Thin wrapper over ``OpcPackage.next_partname(tmpl)``. Saves callers from
   reaching back through ``self._package`` when they only hold the part.


Use-cases unlocked
------------------

F1 is a building block — it doesn't ship any user-visible API on its own. It
is designed to be consumed by:

* F5 (embedded-workbook cloner for chart copy)
* F6 (slide-id manager for insert/reorder/delete)
* `#68` (reorder a slide)
* `#132` (duplicate a slide)
* `#533` (duplicate a shape)
* `#338` (copy chart with linked workbook)
* `#116` (replace picture, preserving its placement)
* `#640`/`#835`/`#879` (delete / reorder / duplicate slide variants)

Each of those features will wire ``PartRelationshipCloner.clone(...)`` into
its own entry point; F1 is the single place the rId-and-parts machinery
lives, so those features can focus on their own semantics (slide-id
bookkeeping, placeholder inheritance, ...).
