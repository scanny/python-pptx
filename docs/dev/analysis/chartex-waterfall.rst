
Waterfall chart (chartex) — design analysis
===========================================

Design companion to issues `#479`_ and `#651`_ (both requesting
"Waterfall Charts"; they are the same ask and will be resolved together).
A *waterfall chart* is one of the seven Office 2016+ "extended" chart
kinds that live under the ``cx:`` namespace rather than the legacy ``c:``
grammar. This note covers the specifics of *authoring* (creating) a
waterfall chart — the XML shape python-pptx must be able to emit, and
the narrow slice of the chartex grammar that waterfall actually
exercises on top of the funnel baseline.

.. _#479: https://github.com/scanny/python-pptx/issues/479
.. _#651: https://github.com/scanny/python-pptx/issues/651
.. _#305: https://github.com/scanny/python-pptx/issues/305
.. _#386: https://github.com/scanny/python-pptx/issues/386
.. _#583: https://github.com/scanny/python-pptx/issues/583


Status — blocked on F4
----------------------

This issue cannot ship until the chartex foundation (F4) is fully
implemented in master. As of this writing master contains:

* **Wave 1 / #386** — chartex *passthrough*. The ``cx:`` namespace is
  registered, the ``application/vnd.ms-office.chartex+xml`` content type
  and the ``.../2014/relationships/chartEx`` relationship type are known,
  and ``GraphicFrame`` reports ``has_chartex`` / returns
  :attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX` for frames that reference a
  chartex part. The part itself round-trips as opaque bytes via the
  generic :class:`Part` fallthrough.
* **Wave 3 / #583 + chartex-foundation.rst** — *design analysis* for the
  full F4 foundation. See :doc:`chartex-foundation`.
* **Wave 6 / #305 + chartex-funnel.rst** — *design analysis* for the
  first concrete chartex authoring target (funnel). See
  :doc:`chartex-funnel`.

What is still missing, and what #479/#651 therefore needs, is the
specialised ``ChartExPart`` plus the ``cx:`` oxml element tree and a
writer family. None of that has landed. Until it does, the library
cannot create a waterfall chart — only preserve one round-tripped from
PowerPoint (which the existing fixture
``features/steps/test_files/cht-chartex.pptx`` already demonstrates — it
*is* a waterfall).

The rest of this note assumes F4 will ship substantially as described in
:doc:`chartex-foundation` and focuses on the increment waterfall adds on
top of both it and the funnel baseline documented in
:doc:`chartex-funnel`.


Scope of this note
------------------

The F4 design note (:doc:`chartex-foundation`) covers the *shared*
infrastructure — part class, namespace, content type, the oxml subpackage
layout, the ``XL_CHART_TYPE.WATERFALL = 119`` enum member, and the
``cx:series/@layoutId`` dispatch. The funnel design note
(:doc:`chartex-funnel`) covers the *minimum* chartex writer shape
(``cx:chartSpace`` / ``cx:chartData`` / ``cx:plotArea`` /
``cx:plotAreaRegion`` / ``cx:series`` with one ``cx:dataId``). This note
is narrower still: it describes only the subset of the chartex grammar
that waterfall adds *on top* of that baseline, so the #479/#651 feature
work has a concrete target.

Specifically:

* the ``cx:series`` shape with ``layoutId="waterfall"``,
* the ``cx:layoutPr/cx:subtotals`` block that marks individual data
  points as *subtotal* bars (the defining visual feature of a waterfall
  chart),
* the ``cx:layoutPr/cx:visibility/@connectorLines`` toggle for the thin
  horizontal lines that link successive bars,
* the axis configuration waterfall requires (unlike funnel, waterfall
  *does* have ``cx:axis`` children),
* the minimum viable authoring API on ``Shapes.add_chart(...)``,
  including a way to flag subtotal categories.


How waterfall differs from funnel
---------------------------------

The funnel analysis calls waterfall out explicitly as the next-simplest
chartex kind. The structural delta is small but not empty:

================================  ===========  ===========
Feature                           Funnel       Waterfall
================================  ===========  ===========
One series per chart              yes          yes
One cat + one val dimension       yes          yes
``cx:series/@layoutId``           ``funnel``   ``waterfall``
``cx:axis`` children              **none**     **two** (cat + val)
``cx:layoutPr`` on series         absent       **present** (subtotals + connectors)
Per-point styling needed for MVP  no           no (colour defaults by sign)
Default ``cx:dataLabels/@pos``    ``ctr``      ``outEnd``
Supports negative values          n/a (rare)   **yes, first-class**
================================  ===========  ===========

In other words, waterfall is "funnel + an axis pair + a ``cx:layoutPr``
block carrying the subtotal indices and the connector-lines toggle".
Everything else (``cx:chartData`` dimension shape, ``cx:dataId`` wiring,
``cx:externalData`` linkage, embedded xlsx authoring) is identical.

The ``outEnd`` data-label default (vs funnel's ``ctr``) reflects
PowerPoint's convention: bars in a waterfall are typically narrow
vertical bars and the numeric label lives above (or below, for negative
bars) the bar rather than centred inside it.


Target XML shape
----------------

Below is the canonical shape PowerPoint emits for a minimal waterfall
chart with five categories, where the last category is a subtotal (e.g.
*Start, Q1, Q2, Q3, End*). The structure is the target for the
feature-level ``ExChartXmlWriter`` for waterfall.

.. highlight:: xml

Package parts
~~~~~~~~~~~~~

Identical to funnel — see :doc:`chartex-funnel` ("Package parts"). The
content-type override, the ``.../2014/relationships/chartEx``
relationship, and the ``p:graphicFrame`` wrapping are indistinguishable
between waterfall and funnel; the ``cx:series/@layoutId`` inside the
part is the only marker.

The chartEx part
~~~~~~~~~~~~~~~~

::

  <cx:chartSpace
      xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <cx:chartData>
      <cx:data id="0">
        <cx:strDim type="cat">
          <cx:f>Sheet1!$A$2:$A$6</cx:f>
          <cx:lvl ptCount="5">
            <cx:pt idx="0">Start</cx:pt>
            <cx:pt idx="1">Q1</cx:pt>
            <cx:pt idx="2">Q2</cx:pt>
            <cx:pt idx="3">Q3</cx:pt>
            <cx:pt idx="4">End</cx:pt>
          </cx:lvl>
        </cx:strDim>
        <cx:numDim type="val">
          <cx:f>Sheet1!$B$2:$B$6</cx:f>
          <cx:lvl ptCount="5" formatCode="General">
            <cx:pt idx="0">100</cx:pt>
            <cx:pt idx="1">25</cx:pt>
            <cx:pt idx="2">-10</cx:pt>
            <cx:pt idx="3">15</cx:pt>
            <cx:pt idx="4">130</cx:pt>
          </cx:lvl>
        </cx:numDim>
      </cx:data>
    </cx:chartData>
    <cx:chart>
      <cx:plotArea>
        <cx:plotAreaRegion>
          <cx:plotSurface/>
          <cx:series layoutId="waterfall" hidden="0" ownerIdx="0">
            <cx:tx><cx:txData><cx:v>Change</cx:v></cx:txData></cx:tx>
            <cx:layoutPr>
              <cx:subtotals>
                <cx:subtotal idx="0"/>
                <cx:subtotal idx="4"/>
              </cx:subtotals>
              <cx:visibility connectorLines="1"/>
            </cx:layoutPr>
            <cx:dataLabels pos="outEnd"/>
            <cx:dataId val="0"/>
          </cx:series>
        </cx:plotAreaRegion>
        <cx:axis id="0">
          <cx:catScaling gapWidth="0.5"/>
          <cx:majorTickMarks type="none"/>
        </cx:axis>
        <cx:axis id="1">
          <cx:valScaling/>
          <cx:majorTickMarks/>
        </cx:axis>
      </cx:plotArea>
    </cx:chart>
    <cx:externalData r:id="rId1" autoUpdate="0">
      <cx:autoUpdate val="0"/>
    </cx:externalData>
  </cx:chartSpace>

Notable observations:

* ``cx:series/@layoutId="waterfall"`` is the chart-kind marker. Beyond
  that attribute value, the interesting waterfall-specific content is
  entirely inside ``cx:layoutPr``.
* ``cx:layoutPr/cx:subtotals`` enumerates the data-point indices that
  PowerPoint should render as *subtotal* (non-floating) bars. A
  subtotal bar is drawn from the zero baseline rather than from the
  running total of the previous bar, and is typically shaded a third
  colour distinct from the positive (default blue) and negative
  (default orange/red) fills. The canonical waterfall has subtotals at
  index 0 (the starting total) and index *n-1* (the final total); any
  intermediate bar can also be flagged.
* ``cx:layoutPr/cx:visibility/@connectorLines`` is the toggle for the
  thin horizontal step-lines that connect the top of each bar to the
  bottom of the next. ``"1"`` (on) is PowerPoint's default — the
  authoring API should emit this only if the user opts out, or emit it
  unconditionally to match PowerPoint-authored output.
* ``cx:axis`` children are required for waterfall: one category-scale
  axis (``@id="0"`` + ``cx:catScaling``) and one value-scale axis
  (``@id="1"`` + ``cx:valScaling``). They are cross-referenced
  positionally — chartex uses integer ``@id`` rather than legacy
  chart's ``c:axId`` pair-of-IDs pattern.
* ``cx:dataLabels/@pos="outEnd"`` is PowerPoint's default for waterfall
  (vs ``"ctr"`` for funnel). Other legal values (``"inEnd"``,
  ``"inBase"``, ``"ctr"``) are also accepted; ``"outEnd"`` is the
  sensible default for the MVP.
* ``cx:dataId val="0"`` wires the series to the ``cx:data id="0"``
  block under ``cx:chartData`` — identical to funnel.
* ``cx:externalData`` points at the embedded
  ``xl/worksheets/sheet1.xml`` packaged under
  ``/ppt/embeddings/Microsoft_Excel_Worksheet%d.xlsx`` — identical
  machinery to funnel and to legacy charts. The F4 note confirms
  :class:`EmbeddedXlsxPart` can be reused unchanged.


The ``cx:plotArea`` subset waterfall uses
-----------------------------------------

Mapping from schema to the subset waterfall needs:

* ``cx:plotAreaRegion`` — **required**, one per chart.
* ``cx:plotSurface`` — **required**, empty element; present in every
  PowerPoint-authored waterfall.
* ``cx:series`` — **required (×1)**; exactly one, with
  ``@layoutId="waterfall"``.
* ``cx:axis`` — **required (×2)** for waterfall. One ``cx:catScaling``
  category axis and one ``cx:valScaling`` value axis. The pair is what
  distinguishes waterfall structurally from funnel / treemap / sunburst
  (which have no axes at all).
* ``cx:plotAreaRegion/cx:plotAreaRegionLayout`` — *optional*. Only
  emitted when the author manually positions the plot-area; the MVP
  authoring API can omit it.

The two axes can be emitted as bare skeletons for the MVP — no tick
configuration, no title, no text-properties. PowerPoint renders
sensible defaults from the embedded data range.


Series element subset
~~~~~~~~~~~~~~~~~~~~~

Within ``cx:series`` for a waterfall, the MVP needs to emit:

* ``@layoutId`` — **required**, literal ``"waterfall"``.
* ``@hidden`` — *optional*, default ``"0"``. PowerPoint writes it;
  omitting is also valid.
* ``@ownerIdx`` — *optional*, default ``"0"``. Series index; ``"0"``
  for the sole series.
* ``cx:tx/cx:txData/cx:v`` — *optional*. Series name (e.g. "Change");
  if omitted, PowerPoint shows "Series 1".
* ``cx:layoutPr`` — **required whenever any subtotal is flagged or
  connector-lines are toggled off**. Can be omitted entirely for the
  degenerate case of *all* bars floating and connector-lines on.
* ``cx:layoutPr/cx:subtotals/cx:subtotal`` — *optional*, zero or more.
  Each carries ``@idx`` naming a 0-indexed data-point to render as a
  subtotal bar.
* ``cx:layoutPr/cx:visibility/@connectorLines`` — *optional*, default
  ``"1"`` (on). Set to ``"0"`` to hide connector lines.
* ``cx:dataLabels`` — *optional*. Default ``@pos="outEnd"`` for
  waterfall.
* ``cx:dataId`` — **required**. ``@val`` matches a ``cx:data/@id``.
* ``cx:spPr`` / ``cx:txPr`` — *optional*. Appearance overrides;
  omitting falls back to theme (positive-bar, negative-bar and
  subtotal-bar colours all come from theme accents).
* ``cx:dataPoint`` — **not emitted for MVP**. Per-category fill
  overrides; not needed to get a working waterfall. Note: individual
  bar colouring by sign is *automatic* — PowerPoint applies positive
  vs negative vs subtotal fills itself based on the value sign and the
  ``cx:subtotal`` markers. No ``cx:dataPoint`` is required to get the
  canonical red/green/grey visual.


Subtotals — the defining waterfall feature
------------------------------------------

A waterfall without any ``cx:subtotal`` entries renders every bar as a
"floating" change bar stacked on the running total. In practice this
is rarely what the author wants — the whole point of a waterfall is the
anchor bars at the beginning and end (and sometimes intermediate
checkpoints) that show cumulative totals rather than changes.

``cx:subtotals`` is therefore the main authoring-API affordance unique
to waterfall. Two reasonable shapes for the input:

* **Option A (list of indices).** ``chart_data.subtotal_indices =
  (0, 4)``. Terse, matches the XML exactly, but requires the caller
  to know zero-indexed positions.
* **Option B (per-category flag on** :class:`CategoryChartData` **).**
  A new ``chart_data.add_category(..., is_subtotal=True)`` kwarg. More
  discoverable; requires extending ``Category`` / ``CategoryChartData``
  with a flag that other chart kinds ignore.

The funnel design note kept :class:`CategoryChartData` unchanged;
waterfall is the first kind that *genuinely* needs to extend it. Option
B is the better long-term fit but Option A is trivially backward
compatible and is sufficient for an MVP. Picking Option A for the MVP
and deferring Option B to a follow-up is the proposed path.

Semantics PowerPoint enforces (the writer must mirror):

* ``@idx`` is **0-indexed** into the category dimension.
* Duplicate ``@idx`` values are silently tolerated (PowerPoint treats
  them as one).
* ``@idx`` values outside ``[0, n_categories)`` are rejected on load
  by PowerPoint with a repair message — the writer should validate.
* ``cx:subtotals`` may be absent *or* present-but-empty; both mean "no
  subtotal bars".


Connector lines
---------------

The step-lines that link the top of each non-subtotal bar to the bottom
of the next bar are controlled by a single boolean:

::

  <cx:layoutPr>
    <cx:visibility connectorLines="1"/>   <!-- default, lines shown -->
    <cx:visibility connectorLines="0"/>   <!-- lines hidden -->
  </cx:layoutPr>

PowerPoint's UI exposes this as the *"Show connector lines"* checkbox
in the series formatting pane. The MVP authoring API can simply omit
``cx:visibility`` entirely (PowerPoint treats absence as on) and defer
a public toggle to a follow-up issue. If a toggle is wanted up front,
a ``chart_data.show_connector_lines = False`` attribute on the input
is the minimal surface.

``cx:visibility`` has no styling sub-elements of its own; line
appearance comes from theme defaults. Customising connector-line
colour / weight would require ``cx:spPr`` children on the series level
(or on a sibling ``cx:connectorLines`` element in richer PowerPoint
output) and is explicitly out of scope for the MVP.


Cached-data structure for waterfall
-----------------------------------

Identical to funnel — see :doc:`chartex-funnel` ("Cached-data structure
for funnel"). One ``cx:data`` block, one ``cx:strDim type="cat"``, one
``cx:numDim type="val"``, matching ``@ptCount`` across both dimensions.

The one semantic difference is that **negative values are first-class**
for waterfall. Funnel charts with negative values render oddly but
don't misauthor; waterfall *requires* negative-value support because
the whole chart type exists to visualise gains and losses side by side.
The writer must not reject negative values (it doesn't need to do
anything special either — ``cx:pt`` text content is just the number's
string representation).


``XL_CHART_TYPE.WATERFALL``
---------------------------

Per F4, ``XL_CHART_TYPE.WATERFALL`` takes the value ``119`` (matching
the Excel VBA ``XlChartType`` constant for waterfall).
``GraphicFrame.chart_type`` reads the first ``cx:series/@layoutId``
and, when it equals ``"waterfall"``, returns
``XL_CHART_TYPE.WATERFALL``.


Authoring API — minimum viable surface
--------------------------------------

Once F4 lands, the public API change this issue needs is narrow::

  chart_data = CategoryChartData()
  chart_data.categories = ["Start", "Q1", "Q2", "Q3", "End"]
  chart_data.add_series("Change", (100, 25, -10, 15, 130))
  chart_data.subtotal_indices = (0, 4)    # Option A — MVP

  chart = slide.shapes.add_chart(
      XL_CHART_TYPE.WATERFALL,
      left, top, width, height,
      chart_data,
  ).chart

``chart_data`` is the existing ``CategoryChartData`` — one category
axis, one series — extended with a new ``subtotal_indices`` attribute
(default ``()``). Internally the dispatch on ``chart_type`` selects the
new ``ExChartPart.new_waterfall(...)`` classmethod rather than the
legacy ``ChartPart.new(...)``; waterfall and funnel share most of the
writer template and differ only in the ``@layoutId`` literal, the
``cx:layoutPr`` emission, and the pair of ``cx:axis`` children.

What the returned ``chart`` exposes is open — a minimal read-only
``ExChart`` proxy sufficient to satisfy the add-chart contract is enough
for the MVP; the full surface can follow incrementally.

Invariants the writer must enforce:

* Exactly one series (``CategoryChartData.add_series`` called once).
* At least one category / value pair.
* Category and value counts match.
* Every element of ``subtotal_indices`` is an int in
  ``[0, n_categories)``. Raise :exc:`ValueError` otherwise.
* ``cx:series/@layoutId`` is the literal ``"waterfall"``.
* Two ``cx:axis`` children — one category-scale (``id="0"``), one
  value-scale (``id="1"``) — always emitted.
* An embedded xlsx is authored alongside and linked via
  ``cx:externalData/@r:id``. Reuses :class:`EmbeddedXlsxPart` from F4.


Out of scope for #479/#651
--------------------------

Explicitly deferred to follow-up issues — not part of the waterfall
MVP:

* **Per-point fill/line overrides** (``cx:dataPoint``) — the legacy
  ``Point.format`` equivalent for chartex. Waterfall colours for
  positive / negative / subtotal bars come from theme accents;
  overriding individual bars is follow-up work.
* **Custom connector-line styling** (``cx:spPr`` on ``cx:layoutPr`` /
  dedicated sub-element). Only the visibility toggle is in scope;
  colour / weight is follow-up work.
* **Custom data label number formats and positions beyond the
  default** — ``cx:dataLabels`` has its own extensive subgrammar.
* **Reading a waterfall's cached category/value arrays and subtotal
  flags** via a public ``Plot.categories`` / ``Series.values`` analogue
  — reading is a separate feature.
* **The** :class:`CategoryChartData` **per-category** ``is_subtotal``
  **kwarg (Option B above).** Deferred in favour of the terse
  ``subtotal_indices`` list for the MVP; can be added compatibly
  later.
* **Bulk chart copy / clone operations** for chartex parts.
* **The other five chartex kinds not yet analysed** (treemap, sunburst,
  histogram / Pareto, box-and-whisker, map). Each is its own issue.


Checklist — blocked until F4 lands
----------------------------------

This issue is blocked. When F4 (per :doc:`chartex-foundation`) is fully
implemented in master, the following checklist is the concrete scope
for #479/#651:

**F4 prerequisites (verify before unblocking):**

- [ ] :class:`ChartExPart` registered for
      ``application/vnd.ms-office.chartex+xml`` in
      :data:`pptx.content_type_to_part_class_map`.
- [ ] ``pptx/oxml/chartex/`` subpackage exists with
      ``CT_ChartExChartSpace``, ``CT_ChartExChartData``,
      ``CT_ChartExChart``, ``CT_ChartExPlotArea``,
      ``CT_ChartExSeries``, ``CT_ChartExExternalData`` — or at least
      the subset used by waterfall (note: waterfall additionally
      exercises ``CT_ChartExAxis``, ``CT_ChartExCatScaling``,
      ``CT_ChartExValScaling``, ``CT_ChartExLayoutPr``,
      ``CT_ChartExSubtotals``, ``CT_ChartExVisibility`` — a wider
      subset than funnel).
- [ ] ``XL_CHART_TYPE.WATERFALL = 119`` defined (replacing the
      ``UNSUPPORTED_CHARTEX`` sentinel for ``layoutId="waterfall"``).
- [ ] ``GraphicFrame.chart_type`` returns ``XL_CHART_TYPE.WATERFALL``
      for a chartex graphic-frame whose first series has
      ``layoutId="waterfall"``.

**#479/#651-specific deliverables:**

- [ ] ``ExChartPart.new_waterfall(chart_data, package)`` classmethod
      emits the target XML shape (see "Target XML shape" above).
- [ ] ``ExChartXmlWriter`` (or equivalent) templated for waterfall —
      builds ``cx:chartSpace`` from a :class:`CategoryChartData` input
      plus a ``subtotal_indices`` tuple.
- [ ] ``CategoryChartData`` grows a ``subtotal_indices`` attribute
      (default ``()``); ignored by non-waterfall chart kinds.
- [ ] ``Shapes.add_chart(XL_CHART_TYPE.WATERFALL, ...)`` dispatches to
      ``ExChartPart.new_waterfall`` rather than ``ChartPart.new``.
- [ ] Embedded xlsx is authored with category names in column A and
      numeric values in column B, starting at row 2. Reuses
      :class:`EmbeddedXlsxPart` with no modification.
- [ ] ``cx:externalData/@r:id`` relationship wired to the embedded
      xlsx part.
- [ ] ``cx:layoutPr/cx:subtotals`` emitted whenever
      ``subtotal_indices`` is non-empty; absent otherwise.
- [ ] Category-scale axis (``cx:axis id="0"`` with ``cx:catScaling``)
      and value-scale axis (``cx:axis id="1"`` with ``cx:valScaling``)
      emitted as bare skeletons.
- [ ] Input validation: category count equals value count; at least
      one category; single series only; every ``subtotal_indices``
      element is an int in ``[0, n_categories)``. Raises
      :exc:`ValueError` otherwise.
- [ ] Round-trip test: author a waterfall via ``add_chart``, save,
      reload, confirm ``chart_type == XL_CHART_TYPE.WATERFALL`` and
      the cached categories / values / subtotal indices survive.
- [ ] Acceptance test (``features/chart-add.feature`` or similar) that
      opens the saved .pptx against the pattern used in the existing
      ``cht-chartex.pptx`` fixture (which is itself a waterfall — it
      can double as a golden sample).
- [ ] Minimal ``ExChart`` / ``ExSeries`` read-proxy sufficient for
      ``add_chart`` to return a non-opaque object. Can be read-only
      for the MVP. Shared with the funnel implementation if #305 has
      landed first.
- [ ] HISTORY entry along the lines of:
      *feat: #479/#651 waterfall chart authoring via*
      ``slide.shapes.add_chart(XL_CHART_TYPE.WATERFALL, ...)``.

**Deliberately deferred (not part of #479/#651):**

- Per-point fill overrides (``cx:dataPoint``) for positive / negative
  / subtotal colours beyond the theme defaults.
- Connector-line styling (colour / weight).
- Non-default data-label positions / formats.
- Read-side introspection of the cached categories / values /
  subtotal flags.
- Per-category ``is_subtotal=True`` kwarg on
  ``CategoryChartData.add_category`` (Option B in "Subtotals" above);
  the terse ``subtotal_indices`` attribute is the MVP surface.


Test-fixture plan
-----------------

Two fixtures are needed:

* **Authored-by-PowerPoint waterfall fixture** — the existing
  ``features/steps/test_files/cht-chartex.pptx`` is already a
  PowerPoint-authored waterfall (see the specimen in
  :doc:`chartex-foundation` and the ``cx:series/@layoutId="waterfall"``
  in its chartEx part). It can double as the round-trip parsing
  fixture. One caveat: that specimen does *not* exercise
  ``cx:layoutPr/cx:subtotals`` or ``cx:visibility/@connectorLines`` —
  it is a four-category waterfall with no subtotals flagged. A richer
  second fixture
  (``features/steps/test_files/cht-chartex-waterfall-subtotals.pptx``)
  with the five-category pattern from "Target XML shape" above is
  proposed to exercise the subtotal and connector-line paths.
* **Hand-crafted minimal XML fixture** — an XML string in
  ``tests/chart/fixtures/chartex_waterfall_basic.xml`` (the same new
  directory proposed by :doc:`chartex-funnel`) containing the
  ``cx:chartSpace`` emitted by ``ExChartPart.new_waterfall(...)`` for
  a canonical input. Golden-file tested against the writer output.
  A second variant
  (``tests/chart/fixtures/chartex_waterfall_subtotals.xml``) covers
  the non-empty ``cx:subtotals`` path.


Related work
------------

* :doc:`chartex-foundation` — F4 foundation design (the blocker for
  this issue).
* :doc:`chartex-funnel` — #305 funnel chart design (the sibling
  chartex-authoring issue; shares most of the writer machinery).
* `#583`_ — chartex parent issue covering all seven kinds.
* `#386`_ — Wave 1 passthrough (chartex parts round-trip as opaque
  bytes). Already in master.
