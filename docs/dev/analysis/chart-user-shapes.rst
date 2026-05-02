.. _chart_user_shapes:

Chart user-shapes (``c:userShapes``) — design analysis
======================================================

Context
-------

Issue #351 asks for access to the annotation shapes that PowerPoint lets
a user draw on top of a chart — arrows, call-outs, text boxes, and any
other DrawingML shape — pinned to the chart's plot area so they track
the chart when it is moved or resized.

These annotations live **outside** the chart's own XML. PowerPoint
stores them in a separate *chart-drawing* part (content-type
``application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml``)
linked from the chart part via a ``chartUserShapes`` relationship. The
chart XML carries a ``c:userShapes`` element whose ``r:id`` attribute
names that relationship; the target part's XML root is
``cdr:userShapes`` and its children are *anchors* wrapping single
DrawingML shapes::

    <c:chartSpace ...>
      ...
      <c:userShapes r:id="rId2"/>
    </c:chartSpace>

    <!-- ppt/charts/drawings/drawing1.xml -->
    <cdr:userShapes xmlns:cdr="..." xmlns:a="..." xmlns:r="...">
      <cdr:relSizeAnchor>
        <cdr:from><cdr:x>0.12</cdr:x><cdr:y>0.08</cdr:y></cdr:from>
        <cdr:to><cdr:x>0.42</cdr:x><cdr:y>0.22</cdr:y></cdr:to>
        <cdr:sp>...(textbox)...</cdr:sp>
      </cdr:relSizeAnchor>
      <cdr:absSizeAnchor>
        <cdr:from><cdr:x>0.6</cdr:x><cdr:y>0.7</cdr:y></cdr:from>
        <cdr:ext cx="914400" cy="228600"/>
        <cdr:cxnSp>...(arrow)...</cdr:cxnSp>
      </cdr:absSizeAnchor>
    </cdr:userShapes>

The grammar is defined in ``spec/ISO-IEC-29500-1/schemas/xsd/dml-chartDrawing.xsd``.


Anchor model
------------

Two anchor flavours are permitted as direct children of
``cdr:userShapes``:

``cdr:relSizeAnchor``
    Two fractional coordinates (``cdr:from`` / ``cdr:to``) in the
    ``[0.0, 1.0]`` plot-area coordinate space. The shape *resizes*
    with the chart.

``cdr:absSizeAnchor``
    A fractional origin (``cdr:from``) plus a fixed EMU extent
    (``cdr:ext``). The shape keeps its size when the chart resizes.

Each anchor wraps exactly one shape-like child — ``cdr:sp``,
``cdr:grpSp``, ``cdr:graphicFrame``, ``cdr:cxnSp``, or ``cdr:pic`` —
chosen from ``EG_ObjectChoices``. The shape grammar is the familiar
DrawingML one (``CT_Shape`` / ``CT_Connector`` / etc.), with the
inner ``CT_NonVisualDrawingProps`` / ``CT_ShapeProperties`` /
``CT_TextBody`` trees unchanged from the slide-tree forms except that
they live under the ``cdr:`` namespace and their ``a:``-qualified
children are still DrawingML-main.


MVP scope (this patch)
----------------------

This patch ships a **read-only scaffold**. What is included:

* Content-type registration of
  :class:`pptx.parts.chartdrawing.ChartDrawingPart` against
  ``CT.DML_CHARTSHAPES`` so user-shapes parts survive round-trip
  (they had previously been loaded as an anonymous ``Part`` with
  opaque bytes).
* The ``cdr`` namespace is added to
  :mod:`pptx.oxml.ns` so downstream code can qualify
  chart-drawing element names via ``qn("cdr:userShapes")``.
* :attr:`pptx.chart.chart.Chart.user_shapes` resolves the
  ``chartUserShapes`` relationship on the chart part and returns the
  :class:`~pptx.parts.chartdrawing.ChartDrawingPart` (or ``None`` if
  absent).
* :attr:`pptx.chart.chart.Chart.has_user_shapes` is a non-destructive
  probe: ``True`` iff the relationship exists and the target part
  carries at least one anchor.
* :meth:`~pptx.parts.chartdrawing.ChartDrawingPart.iter_anchor_elements`
  and :attr:`~pptx.parts.chartdrawing.ChartDrawingPart.anchor_count`
  yield the raw ``lxml`` anchor elements from ``cdr:userShapes``. No
  python-pptx proxy class is returned — callers inspect XML directly,
  round-trip fidelity is preserved byte-for-byte.

What the MVP deliberately does **not** include:

* **Authoring.** There is no public API to add a new annotation shape.
  :meth:`ChartDrawingPart.new` does create an empty part so the code
  path is tested, but there is no :meth:`Chart.add_user_shape` /
  anchor-level proxy / text-frame helper. The chart on import keeps
  whatever user-shapes it had; a chart authored through this library
  has none.
* **Proxy classes** for ``cdr:relSizeAnchor`` / ``cdr:absSizeAnchor`` /
  ``cdr:sp`` / ``cdr:cxnSp`` / ``cdr:pic`` / ``cdr:graphicFrame`` /
  ``cdr:grpSp``. See below for why.
* **Relationship-aware shapes.** A ``cdr:pic`` may carry its own image
  relationship to a ``cdr:blipFill/a:blip/@r:embed`` target; this
  patch does not resolve those rels.


Why authoring is deferred
-------------------------

Authoring a chart user-shape is not a small extension of the existing
chart-XML work — it is most of the slide-shape stack mirrored into
the ``cdr:`` namespace. To do it properly the library would need:

1. A ``cdr``-namespaced parallel of the shape-tree classes in
   :mod:`pptx.shapes` and :mod:`pptx.oxml.shapes` — at minimum a
   ``UserShape`` / ``UserConnector`` / ``UserPicture`` /
   ``UserGraphicFrame`` / ``UserGroupShape`` set, each backed by a
   ``cdr:``-qualified oxml element class. The shape grammar is shared
   with the slide-tree (``CT_Shape`` is defined in ``dml-main.xsd``),
   but the element tag names are ``cdr:sp`` etc., so element-class
   registration must be duplicated.
2. An anchor proxy (``_BaseAnchor`` with ``RelSizeAnchor`` /
   ``AbsSizeAnchor`` concretes) exposing plot-area-relative coordinates
   in a usable form. The fractional coordinates are handy for PowerPoint
   but unintuitive for typical python-pptx callers who think in EMU /
   ``Inches``, so the API would want to convert on the fly against the
   chart's current plot-area extent — except the chart does not know
   its own plot-area extent without a layout pass.
3. A :class:`pptx.shapes.shapetree.UserShapes` collection — a
   ``len()``-able iterable analogous to :class:`Slide.shapes` on the
   slide tree — so calling code has a familiar shape-tree idiom. This
   pulls in shape-id allocation, placeholder handling (moot here —
   user-shapes cannot be placeholders), and the add/remove/iterate
   protocol.
4. Text-frame handling for ``cdr:sp/cdr:txBody``. The ``a:``-namespaced
   text body is the same one :class:`~pptx.text.text.TextFrame` already
   speaks, so this part is mostly reusable — but the *text-body parent*
   (``cdr:sp``) is a different element class from ``p:sp``, so the
   ``PartElementProxy`` layer that hands a :class:`TextFrame` its
   parent-proxy must cope.
5. Relationship-layer plumbing for inserting images inside user-shapes
   (``cdr:pic/cdr:blipFill/a:blip/@r:embed`` → new
   :class:`~pptx.parts.image.ImagePart` relationship on the
   :class:`ChartDrawingPart`).

None of the above is conceptually hard, but the total surface is
comparable to one of the existing shape-module PRs. Shipping it as a
single commit would dwarf the read-only MVP this ticket asked for, and
it would be pointless without also negotiating how users are expected
to *construct* the fractional-coordinate anchors (step 2 above) — a
decision that benefits from real reporter feedback on use cases
(programmatic highlights? text call-outs? vector overlays exported
from other tools?).

For now, callers that need to author annotations can do so through
``lxml`` directly, starting from the raw element returned by
:meth:`ChartDrawingPart.iter_anchor_elements`, and building
``cdr:``-qualified children with :func:`pptx.oxml.ns.qn`. A future
follow-up patch can promote the most-used cases to a proxy API
without breaking callers that already target the raw elements,
because the MVP's raw-lxml affordance is explicitly not a removable
convenience.


Round-trip-fidelity note
------------------------

Before this patch, the chart-drawing part was loaded through the
generic :class:`pptx.opc.package.Part` fallback, which keeps the bytes
verbatim. Registering the dedicated part class does not change that
for unmodified files — :class:`XmlPart` reserialises its parsed tree,
which lxml emits in the same element order as it was read (minus
inter-element whitespace, which PowerPoint does not depend on for any
element below the package level). Files that contained annotation
shapes before this patch will continue to contain them after, bit for
bit under ``diff`` ignoring insignificant whitespace.


Open questions
--------------

* **Coordinate convenience.** Does an eventual authoring API expose
  the fractional-coordinate primitive directly, or does it pair with
  an ``Emu``-based helper that needs the chart's plot-area bounds?
  PowerPoint's layout pass computes the plot-area rectangle itself
  (:class:`~pptx.chart.plot.PlotArea` is not fully modeled here);
  python-pptx would have to either require the caller to supply it
  or emit a ``c:plotArea/c:layout/c:manualLayout`` simultaneously.
* **Interaction with linked images.** PowerPoint permits ``cdr:pic``
  anchors to reference image parts via ``r:link`` (external URI) rather
  than ``r:embed``; the authoring API should decide whether to mirror
  the :class:`~pptx.parts.image.Image` API or expose both link modes
  explicitly.
* **Copy semantics on cross-slide chart clone.**
  :meth:`ChartPart.clone_from` relies on
  :class:`~pptx.opc.package.PartRelationshipCloner` to walk every
  ``r:id`` on the source chartSpace; the ``c:userShapes/@r:id`` is
  already enumerated and cloned as part of that walk (nothing special
  in this patch), but once authoring lands the cloner will need to
  also deep-copy the *inner* rels on the chart-drawing part (embedded
  images), not just the top-level ``chartUserShapes`` rel.
