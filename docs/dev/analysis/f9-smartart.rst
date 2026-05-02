
SmartArt foundation — F9
========================

Foundation item *F9* — the scaffolding required before python-pptx can do
anything meaningful with SmartArt graphics. This note is the design-analysis
companion to issue `#83`_ ("feature set: SmartArt support"). F9 ships the
minimal, non-lossy discovery layer: SmartArt graphic-frames surface in
``slide.shapes``, report ``MSO_SHAPE_TYPE.IGX_GRAPHIC``, expose a
``has_smart_art`` flag, and offer a raw-XML ``SmartArt`` proxy for the four
backing parts. Full read/write/create support for SmartArt is *XL* scope —
four interlinked parts, a catalogue of ~150 layout algorithms, and a
data-transform graph — and is deliberately out of scope for F9.

.. _#83: https://github.com/scanny/python-pptx/issues/83


Why SmartArt is XL scope
------------------------

A SmartArt graphic on a slide is **not one XML part**; it is *four*,
cross-referenced from the slide graphic-frame. Each has a distinct role, a
distinct schema, and a distinct content-type. They must all agree — a missing
or out-of-sync member drops the SmartArt to a plain-shape fallback in
PowerPoint. This is why "partial support" (the existing project label for
#83) is particularly costly here: unlike tables or legacy charts, you can't
ship a minimally-usable subset without touching all four parts.

The four parts, and how the slide points at them:

.. highlight:: xml

On the slide::

  <p:graphicFrame>
    <p:nvGraphicFramePr>…</p:nvGraphicFramePr>
    <p:xfrm>…</p:xfrm>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
        <dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                    r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/>
      </a:graphicData>
    </a:graphic>
  </p:graphicFrame>

The ``a:graphicData/@uri`` discriminator for SmartArt is
``http://schemas.openxmlformats.org/drawingml/2006/diagram`` (stored in
:data:`pptx.spec.GRAPHIC_DATA_URI_SMART_ART`). The single child of
``a:graphicData`` is a ``dgm:relIds`` element whose four rId attributes
resolve against the **slide part's** relationships to the four backing
parts:

================  ================  ================================================================  =====================================================================
Attribute         rel-type suffix   Content-type                                                      Root element / purpose
================  ================  ================================================================  =====================================================================
``r:dm``          diagramData       ``application/vnd.openxmlformats-officedocument.drawingml``       ``dgm:dataModel`` — the *semantic tree*: nodes (``dgm:pt``), their
                                    ``.diagramData+xml``                                              text runs, and connections (``dgm:cxn``). Authoring happens here.
``r:lo``          diagramLayout     ``application/vnd.openxmlformats-officedocument.drawingml``       ``dgm:layoutDef`` — the *layout algorithm program*. Usually
                                    ``.diagramLayout+xml``                                            referenced from the Office-installed catalogue (Hierarchy, Cycle,
                                                                                                      Process, List, Matrix, Pyramid, Relationship, Picture) rather than
                                                                                                      authored from scratch.
``r:cs``          diagramColors     ``application/vnd.openxmlformats-officedocument.drawingml``       ``dgm:colorsDef`` — a *color-variation* binding theme color slots
                                    ``.diagramColors+xml``                                            to a color transform (e.g. ``Colorful - Accent Colors``,
                                                                                                      ``Gradient Range - Accent 1``).
``r:qs``          diagramQuickStyle ``application/vnd.openxmlformats-officedocument.drawingml``       ``dgm:styleDef`` — the *style variation* picking fill / line /
                                    ``.diagramStyle+xml``                                             effect recipes (e.g. ``Simple Fill``, ``Inset``, ``Polished``).
================  ================  ================================================================  =====================================================================

A fifth, optional part is often present:

* **diagramDrawing** — content-type
  ``application/vnd.ms-office.drawingml.diagramDrawing+xml``, relationship
  type ``.../relationships/diagramDrawing``, root element
  ``dsp:drawing``. PowerPoint caches the *rendered output* (the flattened
  ``dsp:sp`` list produced by running the layout over the data) in a
  hidden extension under the diagramData part. This is an AlternateContent
  fallback so older or non-SmartArt-aware consumers can still render
  something. python-pptx does not need to emit the drawing cache to stay
  valid (PowerPoint re-computes it on open), but must preserve it when
  round-tripping.

All six relationship / content-type constants are already declared in
:mod:`pptx.opc.constants` (``CONTENT_TYPE.DML_DIAGRAM_{DATA,LAYOUT,COLORS,
STYLE,DRAWING}`` and ``RELATIONSHIP_TYPE.DIAGRAM_{DATA,LAYOUT,COLORS,
QUICK_STYLE}``). No part class is registered against these content-types,
so the :class:`~pptx.opc.package.PartFactory` fallthrough loads them as
generic :class:`~pptx.opc.package.Part` instances — which is exactly what
F9's MVP needs for round-trip preservation.


What F9 ships (MVP)
-------------------

F9's deliverable is *discoverability and round-trip preservation*, nothing
more. Concretely:

1. **New namespace registration**. ``dgm`` →
   ``http://schemas.openxmlformats.org/drawingml/2006/diagram`` added to
   :data:`pptx.oxml.ns._nsmap`. This lets oxml code and
   cxml-driven tests construct ``dgm:relIds`` elements without ad-hoc
   namespace juggling.

2. **New spec constant**. :data:`pptx.spec.GRAPHIC_DATA_URI_SMART_ART` —
   the value of ``a:graphicData/@uri`` that identifies a SmartArt frame.

3. **New oxml element class**. :class:`~pptx.oxml.shapes.graphfrm.CT_DgmRelIds`
   with four typed attribute descriptors (``dm_rId``, ``lo_rId``,
   ``qs_rId``, ``cs_rId``). :class:`~pptx.oxml.shapes.graphfrm.CT_GraphicalObjectData`
   grows a :attr:`~pptx.oxml.shapes.graphfrm.CT_GraphicalObjectData.dgm_relIds`
   property that returns the child element (or |None| for non-SmartArt
   frames).

4. **Surface on GraphicFrame**. Three new members on
   :class:`~pptx.shapes.graphfrm.GraphicFrame`:

   * :attr:`~pptx.shapes.graphfrm.GraphicFrame.has_smart_art` — |True| for
     graphic-frames whose ``graphicData_uri`` matches
     :data:`~pptx.spec.GRAPHIC_DATA_URI_SMART_ART`. This resolves the
     pre-existing "partial support" where SmartArt frames were discoverable
     as :class:`~pptx.shapes.graphfrm.GraphicFrame` but not
     *identifiable* as SmartArt.
   * :attr:`~pptx.shapes.graphfrm.GraphicFrame.shape_type` — now returns
     :attr:`MSO_SHAPE_TYPE.IGX_GRAPHIC` for SmartArt frames (previously
     fell through to |None|).
   * :attr:`~pptx.shapes.graphfrm.GraphicFrame.smart_art` — returns a
     :class:`~pptx.shapes.graphfrm.SmartArt` proxy. Raises
     :exc:`ValueError` for non-SmartArt frames.

5. **New ``SmartArt`` class**. A thin
   :class:`~pptx.shared.ParentedElementProxy` subclass with four read-only
   raw-bytes accessors:

   * :attr:`~pptx.shapes.graphfrm.SmartArt.data_xml`
   * :attr:`~pptx.shapes.graphfrm.SmartArt.layout_xml`
   * :attr:`~pptx.shapes.graphfrm.SmartArt.colors_xml`
   * :attr:`~pptx.shapes.graphfrm.SmartArt.quick_style_xml`

   Each resolves the corresponding rId off ``dgm:relIds``, fetches the
   related part via :meth:`~pptx.opc.package.Part.related_part`, and
   returns its :attr:`~pptx.opc.package.Part.blob`. Returns |None| when
   ``dgm:relIds`` is absent, when the attribute is missing, or when the
   rId cannot be resolved (mirroring
   :attr:`~pptx.parts.chart.ChartWorkbook.xlsx_part`'s defensive style
   from #490).

6. **Round-trip preservation** is automatic. The four diagram parts have
   no registered part class, so
   :class:`~pptx.opc.package.PartFactory` dispatches to
   :class:`~pptx.opc.package.Part`, which preserves the blob verbatim.
   The ``dgm:relIds`` element inside the slide is preserved as-is because
   nothing in the load/save path rewrites it. The same goes for the
   diagramDrawing AlternateContent fallback inside the diagramData part.

F9 deliberately stops here. The harder work layers on top.


What F9 deliberately does not do
--------------------------------

* **No SmartArt authoring API**. No ``slide.shapes.add_smart_art(layout,
  data)``. Creating a valid SmartArt graphic from scratch requires
  (a) picking a layout (~150 in Office's shipped catalogue), (b) building
  a ``dgm:dataModel`` node tree whose shape satisfies that layout's
  constraints (depth, branching, categories), and (c) running the layout
  algorithm to emit the diagramDrawing cache. (c) in particular is a
  full graph-layout engine; PowerPoint re-computes it on open, so a
  library emitting SmartArt can punt on (c) by omitting the drawing
  part and relying on PowerPoint's "rehydrate on next open" behaviour,
  but this still requires a trustworthy implementation of (a) and (b).

* **No structured SmartArt editing**. No ``SmartArt.nodes``,
  ``SmartArt.add_node``, ``SmartArt.root.text``. The raw-bytes accessors
  let advanced users lxml-wrangle the data model directly; the
  structured API is a feature layer (#83) on top of F9.

* **No specialised ``DiagramPart`` / ``DiagramDataPart`` /
  ``DiagramLayoutPart`` / ``DiagramColorsPart`` / ``DiagramStylePart``
  classes**. F9 relies on :class:`~pptx.opc.package.Part` fallthrough.
  The feature work in #83 will register typed part classes so the
  structured editing API has typed oxml to talk to.

* **No ``dgm:dataModel`` oxml classes**. Emitting these as xmlchemy
  descriptor trees is a substantial undertaking (see §3 of the
  chartex-foundation note for a comparable oxml build-out). F9 adds
  only :class:`~pptx.oxml.shapes.graphfrm.CT_DgmRelIds` — the minimum
  needed to resolve the four rIds from the slide.

* **No layout catalogue**. Shipping a catalogue of Office's built-in
  layouts (so users can write ``add_smart_art("Hierarchy", data)``)
  is a content question as much as a code question — the layouts live
  in Office, not in ECMA-376. A feature-level decision for #83.


Per-subsystem roadmap
---------------------

Breaking down the remaining scope by subsystem — each is a plausible
independent feature issue, ordered by dependency rather than priority:

1. Part classes and content-type dispatch
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Register specialised :class:`~pptx.opc.package.XmlPart` subclasses for
the four diagram content types in :mod:`pptx.__init__`'s
``content_type_to_part_class_map``:

* ``DiagramDataPart`` — content-type
  ``application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml``,
  partname template ``/ppt/diagrams/data%d.xml``.
* ``DiagramLayoutPart`` — ``…diagramLayout+xml``, partname
  ``/ppt/diagrams/layout%d.xml``.
* ``DiagramColorsPart`` — ``…diagramColors+xml``, partname
  ``/ppt/diagrams/colors%d.xml``.
* ``DiagramQuickStylePart`` — ``…diagramStyle+xml``, partname
  ``/ppt/diagrams/quickStyle%d.xml``.

None of these need a ``.new()`` factory at the part tier until the
authoring layer wants one. Until then they exist so ``SmartArt``'s
structured API has typed parts to reach for.

The ``diagramDrawing`` part (``…diagramDrawing+xml``) is PowerPoint's
cache, surfaced via an extension under diagramData — it stays as a
generic :class:`~pptx.opc.package.Part` until there's a reason to
parse its ``dsp:drawing`` tree.


2. Data-model oxml (authoring)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Build xmlchemy descriptors for the ``diagramData`` schema in a new
``pptx/oxml/diagram/`` subpackage. The top-level types, in rough
dependency order, are:

``CT_DataModel`` (``dgm:dataModel``)
    Root of a diagramData part. Contains one ``dgm:ptLst`` (all nodes),
    one ``dgm:cxnLst`` (all connections), ``dgm:bg`` (optional
    background), and ``dgm:whole`` (optional whole-graphic formatting).
    Optionally a ``dgm:extLst`` carrying the drawing cache.

``CT_Pt`` (``dgm:pt``)
    A node. Attributes: ``@modelId`` (a unique id — often a GUID in
    ``{…}`` form, sometimes a decimal for assistant/doc nodes),
    ``@type`` (``doc`` / ``node`` / ``asst`` / ``parTrans`` /
    ``sibTrans`` / ``pres``). Children: ``dgm:prSet`` (presentation
    preferences: phldr text, custom T/L/W/H overrides), ``dgm:spPr``
    (shape properties override), ``dgm:t`` (a ``a:txBody``-shaped text
    body).

``CT_Cxn`` (``dgm:cxn``)
    A connection. Attributes: ``@modelId``, ``@type`` (``parOf`` /
    ``presOf`` / ``presParOf`` / ``unknownRelationship``), ``@srcId``,
    ``@destId``, ``@srcOrd``, ``@destOrd``. ``parOf`` connections
    encode the semantic tree (author-visible); ``presOf`` /
    ``presParOf`` bind semantic nodes to presentation (``pres``-type)
    nodes owned by the layout.

``CT_PrSet`` (``dgm:prSet``)
    Optional presentation overrides on a node: ``@phldrT``
    (placeholder text when empty), ``@custT`` (text edited from
    default), ``@loTypeId`` / ``@loCatId`` (layout this node prefers),
    ``@qsTypeId`` / ``@qsCatId`` (style), ``@csTypeId`` /
    ``@csCatId`` (colors), and the ``@custLinFactNeighborX`` family
    of layout-override knobs (dozens of them).

The authoring API (``SmartArt.root_node``, ``node.add_child(text=…)``,
``node.text``) lives on top of these. The non-obvious invariant is
that user-visible nodes (``type="node"``) are only reachable through
``parOf`` connections; breadth-first traversal from the single
``type="doc"`` node (connected to the root via a ``parOf`` cxn) is how
the authoring layer reconstructs the visible tree.


3. Layout subsystem (selection)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The hard question: *which layout do you want?* The diagramLayout part's
``dgm:layoutDef`` is a small programming language for graph layout —
``dgm:varLst`` (layout variables), ``dgm:alg`` (algorithm: composite,
linear, cycle, hierChild, snake, pyramid, tx, connector), ``dgm:shape``
(shape spec), ``dgm:presOf`` (binding to data), ``dgm:constrLst``
(layout constraints), ``dgm:forEach`` (loops). Authoring a new
layoutDef from scratch is impractical in a library context.

The pragmatic path:

* **Library by reference, not by value**. The feature provides an enum
  (``DIAGRAM_LAYOUT_TYPE`` or similar) that maps to the names of Office's
  installed layouts (``OrganizationChart``, ``HorizontalLabeledHierarchy``,
  ``SimplePyramid``, ``BasicProcess``, ``Cycle``, …). Writing a SmartArt
  picks a layout by *looking up the pre-authored layoutDef XML* vendored
  with the library rather than constructing it. The lookup table lives
  in ``src/pptx/data/diagrams/layouts/`` or similar.
* **Source**: extract the ~150 built-in layoutDefs from a PowerPoint
  install (``%ProgramFiles%/Microsoft Office/root/Office16/SmartArt/
  Layouts/1033/*.xml``) or from a sample ``.pptx`` authored with each
  layout. These are the canonical XML PowerPoint writes.
* **Licensing**: the shipped Office layouts are Microsoft IP. A
  Python-pptx ship either (a) ships a minimal hand-authored subset
  under a permissive license (Hierarchy, Cycle, Process, List — the
  four most common) or (b) ships the machinery and documents how to
  point it at a user-supplied layout XML.

Custom layouts (user-authored ``dgm:layoutDef``) are a power-user
extension once the API exists.


4. Styling subsystem
~~~~~~~~~~~~~~~~~~~~

Analogous to §3 but smaller — ``dgm:styleDef`` (``r:qs``) and
``dgm:colorsDef`` (``r:cs``) are fewer and simpler than layoutDefs.
Office ships ~24 style variations and ~30 color transforms. Same
approach: enum → vendored XML → looked up on assignment.

The two style/color parts can be shared across SmartArt graphics in a
deck (the relationship points at them; multiple frames can ``r:qs`` /
``r:cs`` to the same part), so the feature should deduplicate on
add-to-package.


5. Drawing-cache maintenance
~~~~~~~~~~~~~~~~~~~~~~~~~~~~

``dgm:extLst/dgm:ext[@uri="http://schemas.microsoft.com/office/drawing/2008/diagram"]
/dsp:dataModelExt`` carries a pointer to the separate diagramDrawing
part that holds the flattened ``dsp:sp``/``dsp:grpSp`` list PowerPoint
uses as its AlternateContent fallback. When python-pptx mutates the
data tree, the cache becomes stale. Two acceptable responses:

* **Drop the cache**. Delete the ``dgm:extLst`` ext and the
  diagramDrawing part relationship. PowerPoint rebuilds on next open.
  Simple, correct, slightly less friendly to consumers that don't run
  the layout themselves (e.g. LibreOffice older versions).
* **Recompute the cache**. Requires a real layout engine. Out of
  scope; noted for completeness.

F9 preserves the cache on round-trip because nothing mutates the data
model. The data-model authoring layer (§2) ships cache-drop logic.


6. Tests and fixtures
~~~~~~~~~~~~~~~~~~~~~

A new fixture ``features/steps/test_files/smart-art.pptx`` containing
one slide with one SmartArt graphic (a simple 3-node horizontal
hierarchy is sufficient — that's ~5 KB of diagram XML across four
parts). F9's MVP tests are unit-level only; an acceptance test that
loads this fixture and asserts round-trip-identical bytes is good
insurance for any foundation and is a sensible add-on when the fixture
lands.

Per-subsystem work gets its own fixtures: §2 needs per-node-type
minimal diagrams; §3 needs one fixture per layout family it claims to
support; §4 needs per-style/per-color pairings.


XML specimens
-------------

For reference, a minimal SmartArt slide fragment (elided):

.. highlight:: xml

Slide relationships (in ``slides/_rels/slide1.xml.rels``)::

  <Relationship Id="rId2" Target="../diagrams/data1.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"/>
  <Relationship Id="rId3" Target="../diagrams/layout1.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"/>
  <Relationship Id="rId4" Target="../diagrams/quickStyle1.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle"/>
  <Relationship Id="rId5" Target="../diagrams/colors1.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"/>

Content-type overrides (in ``[Content_Types].xml``)::

  <Override PartName="/ppt/diagrams/data1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>
  <Override PartName="/ppt/diagrams/layout1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/>
  <Override PartName="/ppt/diagrams/quickStyle1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/>
  <Override PartName="/ppt/diagrams/colors1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/>

The diagramData root (highly abbreviated)::

  <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <dgm:ptLst>
      <dgm:pt modelId="{…doc…}" type="doc"/>
      <dgm:pt modelId="{…n1…}">
        <dgm:prSet/>
        <dgm:spPr/>
        <dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Top</a:t></a:r></a:p></dgm:t>
      </dgm:pt>
      <dgm:pt modelId="{…n2…}">…Left child…</dgm:pt>
      <dgm:pt modelId="{…n3…}">…Right child…</dgm:pt>
      <dgm:pt modelId="{…pres-doc…}" type="pres">…</dgm:pt>
      …
    </dgm:ptLst>
    <dgm:cxnLst>
      <dgm:cxn modelId="{…c1…}" type="parOf"
               srcId="{…doc…}" destId="{…n1…}" srcOrd="0" destOrd="0"/>
      <dgm:cxn modelId="{…c2…}" type="parOf"
               srcId="{…n1…}" destId="{…n2…}" srcOrd="0" destOrd="0"/>
      <dgm:cxn modelId="{…c3…}" type="parOf"
               srcId="{…n1…}" destId="{…n3…}" srcOrd="1" destOrd="0"/>
      <dgm:cxn modelId="{…cp1…}" type="presOf"
               srcId="{…pres-doc…}" destId="{…doc…}"/>
      …
    </dgm:cxnLst>
    <dgm:bg/>
    <dgm:whole/>
  </dgm:dataModel>

Note the three coexisting node graphs: **author-visible** (``doc`` +
``node`` points, connected by ``parOf``), **presentation** (``pres``-type
points with their own ``parOf`` tree, owned by the layout), and the
**binding** between them (``presOf`` connections). A feature-layer
authoring API operates on the first graph; the other two are layout
plumbing.


Consumers unlocked
------------------

F9 on its own unlocks:

* Round-trip preservation of SmartArt-containing decks (previously
  decks with SmartArt survived, but the shapes were undifferentiated
  from any other unknown graphic frame).
* Identifying SmartArt shapes via ``GraphicFrame.has_smart_art`` and
  ``shape.shape_type == MSO_SHAPE_TYPE.IGX_GRAPHIC``, enabling
  "skip-SmartArt" or "copy-SmartArt-verbatim" patterns in user code.
* Raw-XML inspection via ``SmartArt.data_xml`` etc., giving power users
  a documented seam for lxml-based custom editing without monkey-patching.

Features that compose on F9:

* `#83`_ — structured SmartArt read/write/create (the XL-scope feature
  this foundation is for).
* Any "copy a SmartArt from deck A to deck B" operation — requires
  cross-part cloning of the four diagram parts, which naturally
  composes with the existing cross-part cloning infrastructure once
  typed part classes exist (§1).


Migration strategy
------------------

F9 is purely additive. No existing API changes semantics; the
``GraphicFrame.shape_type`` change from |None| to
:attr:`MSO_SHAPE_TYPE.IGX_GRAPHIC` on SmartArt frames is a strict
refinement — callers that handle |None| as "unknown graphic frame"
continue to work, and callers that pattern-match on
``MSO_SHAPE_TYPE.IGX_GRAPHIC`` (a member that has existed in the enum
for years) start seeing the value they always expected.
