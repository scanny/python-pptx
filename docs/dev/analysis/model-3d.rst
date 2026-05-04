
Embedded 3D models — read-only passthrough (issue #410)
=======================================================

Issue `#410`_ asks for support of PowerPoint 365's "Insert > 3D Models"
feature — embedded glTF / OBJ / FBX 3D-model assets rendered inline on a
slide with a camera pose, scene lighting, and (optionally) a selected
animation. Shipping full authoring requires a scene graph, camera maths,
and an animation API that is beyond one wave. This note is the design
companion for the MVP: *detection and round-trip passthrough only*.

.. _#410: https://github.com/scanny/python-pptx/issues/410


What's on a slide
-----------------

Producing a minimal sample by inserting a 3D model in PowerPoint 365,
unzipping the resulting ``.pptx``, and inspecting the slide XML reveals
the following structure:

.. highlight:: xml

::

  <p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
                  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <p:nvGraphicFramePr>
      <p:cNvPr id="2" name="3D Model 1"/>
      <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
      <p:nvPr/>
    </p:nvGraphicFramePr>
    <p:xfrm>
      <a:off x="914400" y="914400"/>
      <a:ext cx="3657600" cy="3657600"/>
    </p:xfrm>
    <a:graphic>
      <a:graphicData uri="http://schemas.microsoft.com/office/drawing/2016/12/model3D">
        <am3d:model3D r:embed="rId2" ext="glb">
          <am3d:camera.../>
          <am3d:ambLight.../>
          <am3d:lights/>
          <am3d:extents.../>
          <!-- (optional) <am3d:animations/> -->
        </am3d:model3D>
      </a:graphicData>
    </a:graphic>
  </p:graphicFrame>

The ``a:graphicData/@uri`` discriminator for a 3D model is
``http://schemas.microsoft.com/office/drawing/2016/12/model3D`` (stored
in :data:`pptx.spec.GRAPHIC_DATA_URI_MODEL_3D`). The single child of
``a:graphicData`` is an ``am3d:model3D`` element whose ``r:embed``
attribute resolves against the **slide part's** relationships to a
binary model part (typically a ``.glb`` — binary glTF — but ``.obj``
and ``.fbx`` are also observed).

Relationship example (``ppt/slides/_rels/slide1.xml.rels``)::

  <Relationship Id="rId2"
    Type="http://schemas.microsoft.com/office/2017/06/relationships/model3D"
    Target="../media/model1.glb"/>

Content-type override (``[Content_Types].xml``)::

  <Override PartName="/ppt/media/model1.glb"
    ContentType="model/gltf-binary"/>

The ``am3d:model3D`` children encode the scene parameters PowerPoint
needs to render the mesh inline: camera position / target / field-of-view
(``am3d:camera``), ambient-light colour and intensity
(``am3d:ambLight``), point / directional lights (``am3d:lights``), the
model's axis-aligned bounding box (``am3d:extents``), and — for animated
models — the currently-selected animation clip (``am3d:animations``).
These are **Microsoft extensions**; the ``am3d:`` namespace
(``http://schemas.microsoft.com/office/drawing/2017/model3d``) is not
part of ECMA-376 and is not described by any XSD shipped under
``spec/ISO-IEC-29500-1/``. Structure and attribute names were reverse-
engineered from real PowerPoint output.


What the MVP ships
------------------

Issue #410 delivers *discoverability* and *round-trip preservation*:

1. **New namespace registration**. ``am3d`` →
   ``http://schemas.microsoft.com/office/drawing/2017/model3d`` added to
   :data:`pptx.oxml.ns._nsmap`. This lets oxml code and cxml-driven
   tests construct ``am3d:model3D`` elements without ad-hoc namespace
   juggling.

2. **New spec constant**. :data:`pptx.spec.GRAPHIC_DATA_URI_MODEL_3D` —
   the value of ``a:graphicData/@uri`` that identifies a 3D-model frame.

3. **New oxml element class**. :class:`~pptx.oxml.shapes.graphfrm.CT_Model3D`
   with two typed attribute descriptors (``embed_rId`` / ``ext``).
   :class:`~pptx.oxml.shapes.graphfrm.CT_GraphicalObjectData` grows a
   :attr:`~pptx.oxml.shapes.graphfrm.CT_GraphicalObjectData.model_3d`
   property returning the child element (or |None| for non-3D frames).
   The child's descendants (camera, lights, extents, animations) are
   **not** modeled — they ride along as generic lxml elements and
   serialize verbatim on round-trip.

4. **Surface on GraphicFrame**. Three new members on
   :class:`~pptx.shapes.graphfrm.GraphicFrame`:

   * :attr:`~pptx.shapes.graphfrm.GraphicFrame.has_model_3d` — |True|
     for graphic-frames whose ``graphicData_uri`` matches
     :data:`~pptx.spec.GRAPHIC_DATA_URI_MODEL_3D`.
   * :attr:`~pptx.shapes.graphfrm.GraphicFrame.model_3d_xml` — raw
     Unicode serialization of the ``am3d:model3D`` subtree (including
     lxml-emitted namespace declarations so the fragment is well-formed
     on its own). |None| for non-3D frames.
   * :attr:`~pptx.shapes.graphfrm.GraphicFrame.model_3d` — returns a
     :class:`~pptx.shapes.model3d.Model3D` proxy. Raises
     :exc:`ValueError` for non-3D frames.

5. **New ``Model3D`` class**
   (:mod:`pptx.shapes.model3d`). A thin
   :class:`~pptx.shared.ParentedElementProxy` subclass with three
   read-only accessors:

   * :attr:`~pptx.shapes.model3d.Model3D.embedded_rel_id` — the
     ``r:embed`` rId, typed ``str | None``.
   * :attr:`~pptx.shapes.model3d.Model3D.ext` — file-extension
     discriminator (``"glb"`` / ``"obj"`` / ``"fbx"`` / |None|).
   * :attr:`~pptx.shapes.model3d.Model3D.media_blob` — bytes of the
     referenced embedded part (|None| when the ``r:embed`` attribute is
     missing or when the relationship cannot be resolved).

6. **Round-trip preservation** is automatic. The embedded model part
   has no registered part class, so
   :class:`~pptx.opc.package.PartFactory` dispatches to
   :class:`~pptx.opc.package.Part`, which preserves the blob verbatim.
   The ``am3d:model3D`` element inside the slide is preserved as-is
   because nothing in the load / save path rewrites it.


What the MVP deliberately does not do
-------------------------------------

* **No 3D-model authoring API**. No
  ``slide.shapes.add_model_3d(glb_bytes, camera=...)``. Authoring
  requires (a) registering the Microsoft content-type and relationship,
  (b) hand-rolling a scene-parameter tree (camera, lights, extents,
  animations), and (c) an optional animation-selection API. (b) in
  particular means modeling the full ``am3d:*`` schema — and there is
  no XSD to back-stop it, only observed PowerPoint output. The shape
  of the authoring API is deferred until the community need becomes
  concrete (e.g. "programmatically embed a product glTF from a
  pipeline").

* **No scene-parameter introspection**. No
  ``Model3D.camera`` / ``.lights`` / ``.animations`` / ``.extents``
  proxies. A power user who wants to tweak a scene parameter today
  reaches through to ``shape.element`` and mutates the lxml tree
  directly. Structured accessors layer on top of the MVP once the
  authoring work is scoped.

* **No specialised ``Model3DPart`` class**. The MVP relies on
  :class:`~pptx.opc.package.Part` fallthrough. A typed
  ``Model3DPart`` would be warranted only once the authoring layer
  needs to reason about model format (glTF vs glb vs OBJ) or enforce
  content-type invariants on save.

* **No animation API**. PowerPoint lets a user pick which baked
  animation in the glTF plays — an ``am3d:animations/am3d:selected``
  child of ``am3d:model3D`` records it. The MVP preserves that
  selection verbatim; surfacing it as
  ``Model3D.selected_animation`` is a natural §2 follow-up.


Per-subsystem roadmap
---------------------

If / when the full-authoring feature lands, it splits along these
lines:

1. **Content-type and relationship registration.** Add a typed
   ``Model3DPart`` (content-types ``model/gltf-binary`` for glb,
   ``model/gltf+json`` for glTF, ``model/obj`` for OBJ, and
   ``model/fbx`` for FBX). Relationship type:
   ``http://schemas.microsoft.com/office/2017/06/relationships/model3D``.

2. **Scene-parameter oxml.** Build xmlchemy descriptors for the
   ``am3d:camera`` / ``am3d:ambLight`` / ``am3d:lights`` /
   ``am3d:extents`` / ``am3d:animations`` subtrees. Absent an XSD,
   authority comes from PowerPoint output; attribute names / types
   are observed, not specified.

3. **Authoring API.**
   ``slide.shapes.add_model_3d(model_file, left, top, width=None,
   height=None, camera=None)`` — registers the part, the rel, the
   content-type, and emits a sensible default scene (camera framing
   the extents, one ambient plus one key light). Defaults cribbed
   from PowerPoint's "Insert > 3D Models" output for a similar model.

4. **Animation selection.** ``Model3D.animation_name`` getter / setter.
   Requires parsing the embedded model to enumerate animation clips
   (glTF: top-level ``animations`` array in the JSON chunk). Out of
   scope for anything that doesn't also ship a glTF parser.


Consumers unlocked
------------------

The MVP on its own unlocks:

* Round-trip preservation of 3D-model-containing decks (previously
  such decks would have loaded, but the shapes were undifferentiated
  from any other unknown graphic frame and the ``.glb`` part was
  preserved only incidentally because it was related from a known
  slide part).
* Identifying 3D-model shapes via
  ``GraphicFrame.has_model_3d`` — enabling "skip-3D-model" or
  "extract-3D-models" patterns in user code.
* Raw-XML inspection via ``GraphicFrame.model_3d_xml``.
* Dumping the embedded ``.glb`` / ``.obj`` / ``.fbx`` bytes via
  ``Model3D.media_blob`` — useful for extracting models for external
  rendering or batch conversion.


Migration strategy
------------------

The MVP is purely additive. No existing API changes semantics. A
graphic-frame carrying a 3D model previously reported
``shape_type == None``; that stays |None| (no ``MSO_SHAPE_TYPE`` member
for 3D models yet — the enum is MSO-centric and Microsoft hasn't
published one). Callers that pattern-match on ``shape_type`` continue
to hit their "unknown graphic frame" branch; callers that now want to
recognise 3D models use ``has_model_3d`` instead.
