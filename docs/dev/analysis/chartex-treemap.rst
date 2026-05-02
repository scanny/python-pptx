
Treemap chart (chartex) — design analysis
=========================================

Design companion to issue `#371`_ ("Treemap Charts"). A *treemap chart*
is one of the seven Office 2016+ "extended" chart kinds that live under
the ``cx:`` namespace rather than the legacy ``c:`` grammar. This note
covers the specifics of *authoring* (creating) a treemap chart — the
XML shape python-pptx must be able to emit, and the narrow slice of the
chartex grammar that treemap actually exercises on top of the funnel /
waterfall baseline.

Also closes as duplicate-of-#371: `#944`_ ("TreeMap and ScatterPlot"). The
scatter half of #944 is already shipping — ``Shapes.add_chart`` accepts
all five ``XL_CHART_TYPE.XY_SCATTER*`` members and emits a
``c:scatterChart`` that round-trips cleanly through save + reopen (see
:doc:`cht-xy-chart` for the legacy-``c:`` grammar). The treemap half of
#944 is the same feature #371 tracks, blocked on the same F4 work
described below.

.. _#371: https://github.com/scanny/python-pptx/issues/371
.. _#305: https://github.com/scanny/python-pptx/issues/305
.. _#479: https://github.com/scanny/python-pptx/issues/479
.. _#386: https://github.com/scanny/python-pptx/issues/386
.. _#583: https://github.com/scanny/python-pptx/issues/583
.. _#944: https://github.com/scanny/python-pptx/issues/944


Status — blocked on F4
----------------------

This issue cannot ship until the chartex foundation (F4) is fully
implemented in master. As of this writing master contains:

* **Wave 1 / #386** — chartex *passthrough*. The ``cx:`` namespace is
  registered, the ``application/vnd.ms-office.chartex+xml`` content
  type and the ``.../2014/relationships/chartEx`` relationship type
  are known, and ``GraphicFrame`` reports ``has_chartex`` / returns
  :attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX` for frames that reference
  a chartex part. The part itself round-trips as opaque bytes via the
  generic :class:`Part` fallthrough.
* **Wave 3 / #583 + chartex-foundation.rst** — *design analysis* for
  the full F4 foundation. See :doc:`chartex-foundation`.
* **Wave 6 / #305 + chartex-funnel.rst** — *design analysis* for the
  funnel authoring path. See :doc:`chartex-funnel`.
* **Wave 7 / #479 + chartex-waterfall.rst** — *design analysis* for
  the waterfall authoring path. See :doc:`chartex-waterfall`.

What is still missing, and what #371 therefore needs, is the
specialised ``ChartExPart`` plus the ``cx:`` oxml element tree and a
writer family. None of that has landed. Until it does, the library
cannot create a treemap chart — only preserve one round-tripped from
PowerPoint.

The rest of this note assumes F4 will ship substantially as described
in :doc:`chartex-foundation` and focuses on the increment treemap adds
on top of both it and the funnel / waterfall baseline documented in
:doc:`chartex-funnel` and :doc:`chartex-waterfall`.


Scope of this note
------------------

The F4 design note (:doc:`chartex-foundation`) covers the *shared*
infrastructure — part class, namespace, content type, the oxml
subpackage layout, the ``XL_CHART_TYPE.TREEMAP = 117`` enum member,
and the ``cx:series/@layoutId`` dispatch. The funnel design note
(:doc:`chartex-funnel`) covers the *minimum* chartex writer shape
(``cx:chartSpace`` / ``cx:chartData`` / ``cx:plotArea`` /
``cx:plotAreaRegion`` / ``cx:series`` with one ``cx:dataId``). The
waterfall note (:doc:`chartex-waterfall`) adds axes and
``cx:layoutPr``. This note is narrower still: it describes only the
subset of the chartex grammar that treemap adds *on top* of the funnel
baseline, so the #371 feature work has a concrete target.

Specifically:

* the ``cx:series`` shape with ``layoutId="treemap"``,
* the hierarchical ``cx:chartData`` shape — a ``cx:strDim type="cat"``
  with **multiple** ``cx:lvl`` children describing a parent → child
  (→ grandchild …) category hierarchy,
* the ``cx:layoutPr/cx:parentLabelLayout`` attribute that controls how
  PowerPoint paints labels for non-leaf (parent) tiles,
* the absence of any ``cx:axis`` children (treemap, like funnel, has no
  axes),
* the minimum viable authoring API on ``Shapes.add_chart(...)``,
  including how the caller expresses the category hierarchy.


How treemap differs from funnel and waterfall
---------------------------------------------

Treemap is the first chartex kind in this series that needs a genuinely
*different* ``cx:chartData`` shape. Funnel and waterfall are both flat
single-level (category, value) pairs; treemap's whole reason for being
is the hierarchical "rectangles within rectangles" visual, which
requires the category dimension to carry multiple nesting levels.

======================================  ===========  =============  =============
Feature                                 Funnel       Waterfall      Treemap
======================================  ===========  =============  =============
One series per chart                    yes          yes            yes
One cat + one val dimension             yes          yes            yes
``cx:series/@layoutId``                 ``funnel``   ``waterfall``  ``treemap``
``cx:axis`` children                    **none**     **two**        **none**
``cx:strDim`` ``cx:lvl`` count          **1**        **1**          **1..N** (hierarchy)
``cx:layoutPr`` on series               absent       subtotals +    ``parentLabelLayout``
                                                     connectors     + ``dataLabelVisibility``
Per-point styling needed for MVP        no           no             no
Default ``cx:dataLabels/@pos``          ``ctr``      ``outEnd``     ``ctr``
Negative values meaningful              n/a          yes            no (values are sizes)
======================================  ===========  =============  =============

In other words, treemap is "funnel without axes, but with a multi-level
``cx:strDim`` encoding the tree structure, plus a small ``cx:layoutPr``
block that controls parent-tile label painting". The ``cx:chartData``
wiring (``cx:data id="0"`` plus a ``cx:numDim type="val"``), the
``cx:externalData`` linkage, and the embedded xlsx authoring are
identical to funnel / waterfall.

Because values represent *area* (rectangle size), negative values are
not meaningful in a treemap; PowerPoint accepts them but the visual
result is undefined. The writer should warn / reject negative values
in the MVP.


Target XML shape
----------------

Below is the canonical shape PowerPoint emits for a minimal treemap
chart with two hierarchy levels (e.g. *Region → Country*) and six leaf
values. The structure is the target for the feature-level
``ExChartXmlWriter`` for treemap.

.. highlight:: xml

Package parts
~~~~~~~~~~~~~

Identical to funnel / waterfall — see :doc:`chartex-funnel` ("Package
parts"). The content-type override, the
``.../2014/relationships/chartEx`` relationship, and the
``p:graphicFrame`` wrapping are indistinguishable across all three
chart kinds; the ``cx:series/@layoutId`` inside the part is the only
marker.

The chartEx part
~~~~~~~~~~~~~~~~

::

  <cx:chartSpace
      xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <cx:chartData>
      <cx:data id="0">
        <cx:strDim type="cat">
          <cx:f>Sheet1!$A$2:$A$7</cx:f>
          <cx:lvl ptCount="6">
            <cx:pt idx="0">France</cx:pt>
            <cx:pt idx="1">Germany</cx:pt>
            <cx:pt idx="2">Italy</cx:pt>
            <cx:pt idx="3">Japan</cx:pt>
            <cx:pt idx="4">Korea</cx:pt>
            <cx:pt idx="5">China</cx:pt>
          </cx:lvl>
          <cx:lvl ptCount="6">
            <cx:pt idx="0">EMEA</cx:pt>
            <cx:pt idx="1">EMEA</cx:pt>
            <cx:pt idx="2">EMEA</cx:pt>
            <cx:pt idx="3">APAC</cx:pt>
            <cx:pt idx="4">APAC</cx:pt>
            <cx:pt idx="5">APAC</cx:pt>
          </cx:lvl>
        </cx:strDim>
        <cx:numDim type="val">
          <cx:f>Sheet1!$C$2:$C$7</cx:f>
          <cx:lvl ptCount="6" formatCode="General">
            <cx:pt idx="0">120</cx:pt>
            <cx:pt idx="1">180</cx:pt>
            <cx:pt idx="2">90</cx:pt>
            <cx:pt idx="3">150</cx:pt>
            <cx:pt idx="4">60</cx:pt>
            <cx:pt idx="5">240</cx:pt>
          </cx:lvl>
        </cx:numDim>
      </cx:data>
    </cx:chartData>
    <cx:chart>
      <cx:plotArea>
        <cx:plotAreaRegion>
          <cx:plotSurface/>
          <cx:series layoutId="treemap" hidden="0" ownerIdx="0">
            <cx:tx><cx:txData><cx:v>Sales</cx:v></cx:txData></cx:tx>
            <cx:layoutPr>
              <cx:parentLabelLayout val="banner"/>
            </cx:layoutPr>
            <cx:dataLabels pos="ctr"/>
            <cx:dataId val="0"/>
          </cx:series>
        </cx:plotAreaRegion>
      </cx:plotArea>
    </cx:chart>
    <cx:externalData r:id="rId1" autoUpdate="0">
      <cx:autoUpdate val="0"/>
    </cx:externalData>
  </cx:chartSpace>

Notable observations:

* ``cx:series/@layoutId="treemap"`` is the chart-kind marker. Beyond
  that attribute value, the interesting treemap-specific content is
  split between the *multi-``cx:lvl`` ``cx:strDim``* (which encodes
  the hierarchy) and the ``cx:layoutPr`` block (which tunes the
  visual).
* ``cx:strDim type="cat"`` contains **two or more** ``cx:lvl``
  children, ordered **leaf → root**. The first ``cx:lvl`` holds the
  leaf-level labels (one unique entry per data point); subsequent
  ``cx:lvl`` elements hold the parent labels, **repeated** across all
  children of that parent. The example above encodes
  ``EMEA → {France, Germany, Italy}`` and
  ``APAC → {Japan, Korea, China}`` by repeating ``EMEA`` three times
  and ``APAC`` three times in the second level.
* All ``cx:lvl`` children of the same ``cx:strDim`` share a single
  ``@ptCount`` equal to the leaf-count. The hierarchy is expressed by
  *value repetition* at the parent level, not by a shorter parent
  array with child-range pointers. This matches the way Excel persists
  hierarchical pivot labels and mirrors the way ``cx:f`` references a
  multi-column spreadsheet range.
* ``cx:numDim type="val"`` remains single-level — one numeric value
  per leaf, ``@ptCount`` identical to the leaf count of the
  ``cx:strDim``. Parent tile sizes are *derived* by PowerPoint from
  the sum of child leaves; they are **not** persisted.
* ``cx:layoutPr/cx:parentLabelLayout`` carries a single ``@val``
  attribute that selects the label painting mode for non-leaf
  rectangles (see the *"Parent-label layout"* section below).
  PowerPoint's default is ``"banner"`` (a thin titled strip across
  the top of each parent rectangle).
* ``cx:dataLabels/@pos="ctr"`` is PowerPoint's default for treemap
  (vs ``"outEnd"`` for waterfall). Labels render centred inside each
  leaf tile.
* ``cx:dataId val="0"`` wires the series to the ``cx:data id="0"``
  block under ``cx:chartData`` — identical to funnel / waterfall.
* ``cx:externalData`` points at the embedded
  ``xl/worksheets/sheet1.xml`` packaged under
  ``/ppt/embeddings/Microsoft_Excel_Worksheet%d.xlsx`` — identical
  machinery to funnel / waterfall. :class:`EmbeddedXlsxPart` is
  reused unchanged.
* **No** ``cx:axis`` children. Like funnel and sunburst, treemap has
  no axes — tile sizes and labels derive directly from the series
  data.


The ``cx:plotArea`` subset treemap uses
---------------------------------------

Mapping from schema to the subset treemap needs:

* ``cx:plotAreaRegion`` — **required**, one per chart.
* ``cx:plotSurface`` — **required**, empty element; present in every
  PowerPoint-authored treemap.
* ``cx:series`` — **required (×1)**; exactly one, with
  ``@layoutId="treemap"``.
* ``cx:axis`` — **not used** for treemap. Matches funnel; distinguishes
  treemap from waterfall / histogram / Pareto / box-and-whisker.
* ``cx:plotAreaRegion/cx:plotAreaRegionLayout`` — *optional*. Only
  emitted when the author manually positions the plot-area; the MVP
  authoring API can omit it.


Series element subset
~~~~~~~~~~~~~~~~~~~~~

Within ``cx:series`` for a treemap, the MVP needs to emit:

* ``@layoutId`` — **required**, literal ``"treemap"``.
* ``@hidden`` — *optional*, default ``"0"``. PowerPoint writes it;
  omitting is also valid.
* ``@ownerIdx`` — *optional*, default ``"0"``. Series index; ``"0"``
  for the sole series.
* ``cx:tx/cx:txData/cx:v`` — *optional*. Series name (e.g. "Sales");
  if omitted, PowerPoint shows "Series 1".
* ``cx:layoutPr`` — **required whenever a non-default
  ``parentLabelLayout`` is wanted or per-parent label visibility is
  tuned**. Can be omitted entirely; PowerPoint then renders the
  ``"banner"`` parent-label default and shows labels for all tiles.
* ``cx:layoutPr/cx:parentLabelLayout`` — *optional*, ``@val`` is one
  of ``"none"`` / ``"banner"`` / ``"overlapping"`` (see next section).
  Default ``"banner"``.
* ``cx:dataLabels`` — *optional*. Default ``@pos="ctr"`` for treemap.
  ``cx:dataLabels/cx:visibility/@seriesName`` / ``@categoryName`` /
  ``@value`` sub-attributes toggle individual label fields and are
  the main affordance callers reach for after the MVP (showing /
  hiding the leaf category vs. the numeric value inside each tile).
* ``cx:dataId`` — **required**. ``@val`` matches a ``cx:data/@id``.
* ``cx:spPr`` / ``cx:txPr`` — *optional*. Appearance overrides;
  omitting falls back to theme.
* ``cx:dataPoint`` — **not emitted for MVP**. Per-tile fill overrides;
  PowerPoint automatically assigns distinct theme accents to each
  top-level parent tile and shades its children accordingly. No
  ``cx:dataPoint`` is required to get the canonical multi-colour
  visual.


Hierarchical category data
--------------------------

The *defining* authoring challenge for treemap is expressing the
category hierarchy.

Level ordering
~~~~~~~~~~~~~~

``cx:lvl`` children inside ``cx:strDim type="cat"`` are ordered
**leaf-first**. This is counter-intuitive — prose-style descriptions
("EMEA contains France, Germany, Italy") read root-first, but the XML
lists France/Germany/Italy before EMEA. The rationale is that the
*primary* category dimension (the one the ``cx:numDim`` array aligns
with point-for-point via ``cx:pt/@idx``) is the leaf level, and
parent levels are additional metadata layered on top.

All ``cx:lvl`` children share identical ``@ptCount``. The parent
levels repeat their label once per child leaf, so the arrays line up
by index rather than by range.

Uniform depth required
~~~~~~~~~~~~~~~~~~~~~~

PowerPoint treats the hierarchy as a *uniform-depth* tree: every leaf
must live at the same depth. The writer cannot emit a two-level
hierarchy where some leaves sit directly under the root and others
sit one level deeper. Callers who need a "ragged" hierarchy must
either flatten (promote all leaves to the deepest level and duplicate
intermediate labels) or add a synthetic parent to equalise depth.

For the MVP, the writer rejects mixed-depth input with a clear
:exc:`ValueError`. Supporting ragged trees is explicit follow-up
work.

Depth limit
~~~~~~~~~~~

Microsoft's documentation does not publish a hard depth limit; PowerPoint's
UI offers three levels via the ribbon. In practice treemaps beyond three
or four levels become visually useless. The writer imposes no depth
limit of its own but flags very deep inputs as warnings.

Formula cells
~~~~~~~~~~~~~

For a single-level treemap (no parents; equivalent to a flat bar
chart) ``cx:f`` is ``Sheet1!$A$2:$A$N`` just like funnel / waterfall.
For a multi-level treemap the writer authors **multiple columns** in
the embedded xlsx — one column per hierarchy level, leaf on the left
— and ``cx:f`` references the *leaf* column. The additional columns
do not need their own ``cx:f`` attributes; the hierarchy is reified
entirely in the cached ``cx:lvl`` elements. The numeric column shifts
rightward to accommodate (column C in the two-level example above
rather than column B).


Parent-label layout
-------------------

``cx:layoutPr/cx:parentLabelLayout/@val`` selects how PowerPoint paints
labels for **non-leaf** (parent) rectangles:

* ``"banner"`` — *default*. A thin titled strip across the top of each
  parent rectangle carries the parent label in a bold font; children
  render in the remaining area below.
* ``"overlapping"`` — the parent label is drawn *on top of* the
  parent rectangle, overlapping its children. Useful when the
  banner would be too narrow to read; sacrifices some child-tile
  clarity.
* ``"none"`` — parent labels are suppressed entirely. Only leaf
  labels are drawn. Suitable when the parent grouping is visually
  obvious from the color grouping and a caption would be noise.

PowerPoint's UI exposes these as radio buttons in the series
formatting pane.

Per-data-label visibility (sibling ``cx:dataLabels/cx:visibility``
attributes ``@seriesName``, ``@categoryName``, ``@value``) toggles the
individual fields concatenated into each leaf label. The MVP leaves
the visibility defaults in place (category name + value, which is what
PowerPoint emits without a ``cx:visibility`` override). Exposing the
visibility knob is explicit follow-up work.


Cached-data structure summary
-----------------------------

One ``cx:data`` block. One ``cx:strDim type="cat"`` with one-or-more
``cx:lvl`` children (leaf-first ordering, uniform ``@ptCount``). One
``cx:numDim type="val"`` with a single ``cx:lvl`` of leaf values,
matching ``@ptCount``.

No negative values (reject in the writer). No empty / NaN leaves for
the MVP (treat as zero-sized tiles; PowerPoint would render them
degenerately).


``XL_CHART_TYPE.TREEMAP``
-------------------------

Per F4, ``XL_CHART_TYPE.TREEMAP`` takes the value ``117`` (matching
the Excel VBA ``XlChartType`` constant for treemap).
``GraphicFrame.chart_type`` reads the first ``cx:series/@layoutId``
and, when it equals ``"treemap"``, returns
``XL_CHART_TYPE.TREEMAP``.


Authoring API — minimum viable surface
--------------------------------------

The hierarchy is the novel input shape. Two reasonable caller
surfaces:

* **Option A (multi-column categories).** Extend
  :class:`~pptx.chart.data.CategoryChartData` to accept
  ``categories`` that are tuples — one tuple per leaf, one tuple
  element per hierarchy level, ordered root-first for caller
  ergonomics. Example::

    chart_data = CategoryChartData()
    chart_data.categories = [
        ("EMEA", "France"),
        ("EMEA", "Germany"),
        ("EMEA", "Italy"),
        ("APAC", "Japan"),
        ("APAC", "Korea"),
        ("APAC", "China"),
    ]
    chart_data.add_series("Sales", (120, 180, 90, 150, 60, 240))

  The writer flips the tuples to leaf-first order when emitting
  ``cx:lvl`` elements. Terse; fits the existing
  ``CategoryChartData`` shape; scalar-string categories
  (``"France"``) continue to work as a single-level degenerate
  case.

* **Option B (nested ``add_category`` / child API).** A richer
  tree-building API — ``chart_data.add_category("EMEA").add_leaf("France", 120)``
  style. More discoverable for deep hierarchies; requires more new
  public surface and departs from the flat-series pattern the rest
  of the chart API uses.

Option A is the MVP. Option B can be added compatibly later as
sugar over the same underlying data structure.

The public API change this issue needs is then::

  chart_data = CategoryChartData()
  chart_data.categories = [
      ("EMEA", "France"),
      ("EMEA", "Germany"),
      ("EMEA", "Italy"),
      ("APAC", "Japan"),
      ("APAC", "Korea"),
      ("APAC", "China"),
  ]
  chart_data.add_series("Sales", (120, 180, 90, 150, 60, 240))

  chart = slide.shapes.add_chart(
      XL_CHART_TYPE.TREEMAP,
      left, top, width, height,
      chart_data,
  ).chart

Internally the dispatch on ``chart_type`` selects the new
``ExChartPart.new_treemap(...)`` classmethod rather than the legacy
``ChartPart.new(...)``. Treemap shares most of the writer template
with funnel (no axes, one ``cx:dataId``) and differs only in the
``@layoutId`` literal, the multi-``cx:lvl`` ``cx:strDim`` emission,
and the optional ``cx:layoutPr/cx:parentLabelLayout`` block.

What the returned ``chart`` exposes is open — a minimal read-only
``ExChart`` proxy sufficient to satisfy the add-chart contract is
enough for the MVP; the full surface can follow incrementally.

Invariants the writer must enforce:

* Exactly one series (``CategoryChartData.add_series`` called once).
* At least one leaf / value pair.
* Leaf and value counts match.
* Every category entry has the same tuple length (uniform-depth
  hierarchy); scalar-string entries are allowed as a single-level
  shorthand.
* No negative values.
* ``cx:series/@layoutId`` is the literal ``"treemap"``.
* An embedded xlsx is authored alongside and linked via
  ``cx:externalData/@r:id``. Reuses :class:`EmbeddedXlsxPart` from
  F4.


Out of scope for #371
---------------------

Explicitly deferred to follow-up issues — not part of the treemap
MVP:

* **Per-tile fill/line overrides** (``cx:dataPoint``) — the legacy
  ``Point.format`` equivalent for chartex. Treemap tiles pick up
  distinct theme accents per top-level parent automatically;
  per-tile overrides are follow-up work.
* **``parentLabelLayout`` selector on the input.** MVP emits the
  default ``"banner"`` (or omits ``cx:layoutPr`` entirely, which
  PowerPoint renders as banner). ``"overlapping"`` / ``"none"`` can
  be exposed via a ``chart_data.parent_label_layout`` attribute
  later.
* **Data-label field visibility toggles** (``cx:dataLabels/cx:visibility``
  ``@categoryName`` / ``@value`` / ``@seriesName``). MVP accepts
  PowerPoint's defaults.
* **Ragged (mixed-depth) hierarchies.** Rejected by the MVP
  writer; requires a separate design pass on the caller-facing
  shape.
* **Reading a treemap's cached categories / values / hierarchy**
  via a public ``Plot.categories`` / ``Series.values`` analogue.
  Reading is a separate feature.
* **Bulk chart copy / clone operations** for chartex parts.
* **The other four chartex kinds not yet analysed** (sunburst,
  histogram / Pareto, box-and-whisker, map). Sunburst is the
  closest sibling — it shares treemap's hierarchical
  ``cx:strDim`` shape almost verbatim and differs only in the
  ``@layoutId`` literal and the visual. A successful #371
  implementation is ~80% of the sunburst writer.


Checklist — blocked until F4 lands
----------------------------------

This issue is blocked. When F4 (per :doc:`chartex-foundation`) is
fully implemented in master, the following checklist is the concrete
scope for #371:

**F4 prerequisites (verify before unblocking):**

- [ ] :class:`ChartExPart` registered for
      ``application/vnd.ms-office.chartex+xml`` in
      :data:`pptx.content_type_to_part_class_map`.
- [ ] ``pptx/oxml/chartex/`` subpackage exists with
      ``CT_ChartExChartSpace``, ``CT_ChartExChartData``,
      ``CT_ChartExChart``, ``CT_ChartExPlotArea``,
      ``CT_ChartExSeries``, ``CT_ChartExExternalData`` — or at
      least the subset used by treemap (note: treemap additionally
      exercises ``CT_ChartExStrDim`` with *multi-``cx:lvl``* support,
      ``CT_ChartExLayoutPr``, and
      ``CT_ChartExParentLabelLayout`` — a wider subset than funnel
      but narrower than waterfall on the axis side).
- [ ] ``XL_CHART_TYPE.TREEMAP = 117`` defined (replacing the
      ``UNSUPPORTED_CHARTEX`` sentinel for ``layoutId="treemap"``).
- [ ] ``GraphicFrame.chart_type`` returns ``XL_CHART_TYPE.TREEMAP``
      for a chartex graphic-frame whose first series has
      ``layoutId="treemap"``.

**#371-specific deliverables:**

- [ ] ``ExChartPart.new_treemap(chart_data, package)`` classmethod
      emits the target XML shape (see "Target XML shape" above).
- [ ] ``ExChartXmlWriter`` (or equivalent) templated for treemap —
      builds ``cx:chartSpace`` from a :class:`CategoryChartData`
      input whose categories are tuples describing the hierarchy.
- [ ] ``CategoryChartData`` accepts tuple-valued categories
      (root-first) as a valid shape; scalar strings continue to
      work for single-level charts. Ignored / rejected by chart
      kinds that don't support hierarchy.
- [ ] ``Shapes.add_chart(XL_CHART_TYPE.TREEMAP, ...)`` dispatches
      to ``ExChartPart.new_treemap`` rather than ``ChartPart.new``.
- [ ] Embedded xlsx is authored with one column per hierarchy
      level (root in column A, leaf in column *k*) and numeric
      values in column *k+1*, starting at row 2. Reuses
      :class:`EmbeddedXlsxPart` with no modification.
- [ ] ``cx:externalData/@r:id`` relationship wired to the embedded
      xlsx part.
- [ ] ``cx:strDim type="cat"`` emitted with *n* ``cx:lvl`` children
      (one per hierarchy level, leaf-first). All share identical
      ``@ptCount``. Parent-level ``cx:pt`` entries repeat the
      parent label across all child indices.
- [ ] ``cx:layoutPr/cx:parentLabelLayout`` emitted only when the
      caller overrides the default ``"banner"`` — MVP always emits
      the default or omits entirely.
- [ ] No ``cx:axis`` children emitted (treemap has no axes).
- [ ] Input validation: leaf count equals value count; at least
      one leaf; single series only; every category tuple has the
      same length; no negative values. Raises :exc:`ValueError`
      otherwise.
- [ ] Round-trip test: author a treemap via ``add_chart``, save,
      reload, confirm ``chart_type == XL_CHART_TYPE.TREEMAP`` and
      the cached categories (including hierarchy) / values
      survive.
- [ ] Acceptance test (``features/chart-add.feature`` or similar)
      that opens the saved .pptx against the pattern used in the
      existing ``cht-chartex.pptx`` fixture.
- [ ] Minimal ``ExChart`` / ``ExSeries`` read-proxy sufficient for
      ``add_chart`` to return a non-opaque object. Can be
      read-only for the MVP. Shared with the funnel / waterfall
      implementations if #305 / #479 have landed first.
- [ ] HISTORY entry along the lines of:
      *feat: #371 treemap chart authoring via*
      ``slide.shapes.add_chart(XL_CHART_TYPE.TREEMAP, ...)``.

**Deliberately deferred (not part of #371):**

- Per-tile fill overrides (``cx:dataPoint``).
- ``parentLabelLayout`` mode selection (non-default).
- Data-label field visibility toggles.
- Ragged (mixed-depth) hierarchies.
- Read-side introspection of the cached categories / hierarchy /
  values.


Test-fixture plan
-----------------

Two fixtures are needed:

* **Authored-by-PowerPoint treemap fixture** — a .pptx containing a
  minimal two-level treemap chart, saved by PowerPoint 2019+ (2016
  will do — the format hasn't changed). Used for round-trip parsing
  tests and as the golden reference for the XML shape the writer
  must produce. Proposed location:
  ``features/steps/test_files/cht-chartex-treemap.pptx``. A
  single-level variant (flat treemap, equivalent to a nicer-looking
  bar chart) may also be useful as a minimum-complexity specimen.
* **Hand-crafted minimal XML fixture** — an XML string in
  ``tests/chart/fixtures/chartex_treemap_basic.xml`` (the same new
  directory proposed by :doc:`chartex-funnel`) containing the
  ``cx:chartSpace`` emitted by ``ExChartPart.new_treemap(...)`` for
  a canonical two-level input. Golden-file tested against the
  writer output. A single-level variant
  (``tests/chart/fixtures/chartex_treemap_flat.xml``) covers the
  degenerate one-``cx:lvl`` path.


Related work
------------

* :doc:`chartex-foundation` — F4 foundation design (the blocker
  for this issue).
* :doc:`chartex-funnel` — #305 funnel chart design (the sibling
  no-axis chartex-authoring issue; shares most of the writer
  machinery).
* :doc:`chartex-waterfall` — #479/#651 waterfall chart design
  (the sibling axis-bearing chartex-authoring issue).
* `#583`_ — chartex parent issue covering all seven kinds.
* `#386`_ — Wave 1 passthrough (chartex parts round-trip as
  opaque bytes). Already in master.
