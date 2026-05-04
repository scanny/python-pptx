
Box-and-whisker chart (chartex) — design analysis
==================================================

Design companion to issue `#1047`_. A *box-and-whisker* (also called
"box plot") chart is one of the seven Office 2016+ "extended" chart
kinds that live under the ``cx:`` namespace rather than the legacy
``c:`` grammar. This note covers the *authoring* target for a minimum
viable box-and-whisker writer, the XML shape python-pptx must be able to
emit, and the narrow slice of the chartex grammar that box-and-whisker
actually exercises on top of the funnel / waterfall baselines.

.. _#1047: https://github.com/scanny/python-pptx/issues/1047
.. _#386: https://github.com/scanny/python-pptx/issues/386
.. _#583: https://github.com/scanny/python-pptx/issues/583


Status — passthrough MVP shipped; authoring blocked on F4
---------------------------------------------------------

Issue `#1047`_ is the parent request to support box-and-whisker charts.
The **passthrough MVP** is delivered now:

* :attr:`GraphicFrame.has_chartex` is |True| for a box-and-whisker
  graphic-frame (already true for any chartex kind since Wave 1 / #386).
* :attr:`GraphicFrame.chart_type` returns
  :attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX`.
* :attr:`GraphicFrame.chartex_type` (added with this issue) returns the
  raw ``cx:series/@layoutId`` string — for a box-and-whisker that value
  is ``"boxWhisker"``. This lets callers distinguish the chart kind
  without parsing the chartex part themselves. ``chartex_type`` returns
  one of ``"boxWhisker"``, ``"funnel"``, ``"treemap"``, ``"sunburst"``,
  ``"waterfall"``, ``"clusteredColumn"`` (histogram), ``"paretoLine"``
  (Pareto overlay), ``"regionMap"``, or |None| when the graphic-frame is
  not chartex or the chartex part is unresolvable.
* The ``cx:chartSpace`` part and its inline box-and-whisker body are
  preserved byte-identical on save / reload through the generic
  :class:`Part` fallthrough (verified in
  ``tests/test_issue_1047_box_whisker_passthrough.py``).

What is **not** shipped yet, and what structured authoring of
box-and-whisker therefore needs, is the specialised ``ChartExPart`` plus
the ``cx:`` oxml element tree and a writer family. That work is tracked
by the F4 chartex foundation (see :doc:`chartex-foundation`). Until F4
lands, the library cannot *create* a box-and-whisker chart — only
*detect and preserve* one round-tripped from PowerPoint.

The rest of this note assumes F4 will ship substantially as described in
:doc:`chartex-foundation` and focuses on the increment that
box-and-whisker adds on top of the funnel baseline documented in
:doc:`chartex-funnel`.


Scope of this note
------------------

The F4 design note covers the shared infrastructure — part class,
namespace, content type, the oxml subpackage layout, the
``XL_CHART_TYPE.BOX_WHISKER = 121`` enum member, and the
``cx:series/@layoutId`` dispatch. The funnel design note covers the
*minimum* chartex writer shape (``cx:chartSpace`` / ``cx:chartData`` /
``cx:plotArea`` / ``cx:plotAreaRegion`` / ``cx:series`` with one
``cx:dataId``). This note is narrower still: it describes the subset of
the chartex grammar that box-and-whisker adds *on top* of that baseline.

Specifically:

* the ``cx:series`` shape with ``layoutId="boxWhisker"``;
* the ``cx:layoutPr/cx:statistics`` block that selects the quartile
  computation method (``inclusive`` / ``exclusive``);
* the ``cx:layoutPr/cx:visibility`` toggles for the mean marker, mean
  line, non-outlier points, and outlier points — the four visibility
  bits that define a box-and-whisker's look-and-feel;
* the pair of axes (category + value) box-and-whisker requires;
* the minimum viable authoring API on ``Shapes.add_chart(...)``.


How box-and-whisker differs from funnel and waterfall
-----------------------------------------------------

==================================  ==========  ==========  ==============
Feature                             Funnel      Waterfall   Box-Whisker
==================================  ==========  ==========  ==============
One series per chart                yes         yes         yes
``cx:series/@layoutId``             ``funnel``  ``...fall`` ``boxWhisker``
``cx:axis`` children                **none**    **two**     **two**
``cx:layoutPr`` on series           absent      present     **present**
Quartile method selection           n/a         n/a         **yes**
Mean / outlier visibility flags     n/a         n/a         **yes (4)**
Supports negative values            n/a         yes         **yes**
Input shape                         1 cat + val 1 cat + val grouped obs
==================================  ==========  ==========  ==============

Box-and-whisker is structurally closest to waterfall — both carry a
``cx:layoutPr`` block on the series and both require the two-axis
scaffold. Waterfall's ``cx:layoutPr`` wraps ``cx:subtotals`` +
``cx:visibility/@connectorLines``; box-and-whisker's ``cx:layoutPr``
wraps ``cx:statistics`` + ``cx:visibility`` with four boolean toggles.
The chart-kind boundary is the ``@layoutId`` literal on ``cx:series``.

The biggest *semantic* difference is the **input shape**. Funnel and
waterfall consume one category + one value column; box-and-whisker
consumes a *grouped* dataset — typically a category column naming the
*group* (e.g. "A", "A", "A", "B", "B", "B") and a value column carrying
the observations. PowerPoint computes min/Q1/median/Q3/max for each
group at render time; the writer does not need to precompute them.


Target XML shape
----------------

Below is the canonical shape PowerPoint emits for a minimal
box-and-whisker chart with two groups ("A", "B") of two observations
each. The structure is the target for the feature-level
``ExChartXmlWriter`` for box-and-whisker.

.. highlight:: xml

Package parts
~~~~~~~~~~~~~

Identical to funnel / waterfall — see :doc:`chartex-funnel` ("Package
parts"). The content-type override, the
``.../2014/relationships/chartEx`` relationship, and the
``p:graphicFrame`` wrapping are indistinguishable between the chartex
kinds; the ``cx:series/@layoutId`` inside the part is the only marker.

The chartEx part
~~~~~~~~~~~~~~~~

::

  <cx:chartSpace
      xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <cx:chartData>
      <cx:data id="0">
        <cx:strDim type="cat">
          <cx:f>Sheet1!$A$2:$A$5</cx:f>
          <cx:lvl ptCount="4">
            <cx:pt idx="0">A</cx:pt>
            <cx:pt idx="1">A</cx:pt>
            <cx:pt idx="2">B</cx:pt>
            <cx:pt idx="3">B</cx:pt>
          </cx:lvl>
        </cx:strDim>
        <cx:numDim type="val">
          <cx:f>Sheet1!$B$2:$B$5</cx:f>
          <cx:lvl ptCount="4" formatCode="General">
            <cx:pt idx="0">10</cx:pt>
            <cx:pt idx="1">12</cx:pt>
            <cx:pt idx="2">7</cx:pt>
            <cx:pt idx="3">9</cx:pt>
          </cx:lvl>
        </cx:numDim>
      </cx:data>
    </cx:chartData>
    <cx:chart>
      <cx:plotArea>
        <cx:plotAreaRegion>
          <cx:plotSurface/>
          <cx:series layoutId="boxWhisker" hidden="0" ownerIdx="0">
            <cx:tx><cx:txData><cx:v>Values</cx:v></cx:txData></cx:tx>
            <cx:layoutPr>
              <cx:statistics quartileMethod="exclusive"/>
              <cx:visibility meanMarker="1" meanLine="0"
                             nonoutliers="0" outliers="1"/>
            </cx:layoutPr>
            <cx:dataId val="0"/>
          </cx:series>
        </cx:plotAreaRegion>
        <cx:axis id="0">
          <cx:catScaling gapWidth="0.5"/>
        </cx:axis>
        <cx:axis id="1">
          <cx:valScaling/>
        </cx:axis>
      </cx:plotArea>
    </cx:chart>
    <cx:externalData r:id="rId1" autoUpdate="0">
      <cx:autoUpdate val="0"/>
    </cx:externalData>
  </cx:chartSpace>

Notable observations:

* ``cx:series/@layoutId="boxWhisker"`` is the chart-kind marker.
* ``cx:layoutPr/cx:statistics/@quartileMethod`` selects the quartile
  computation. PowerPoint supports ``"inclusive"`` (the default; Q1/Q3
  include the median when computing quartiles) and ``"exclusive"``
  (Q1/Q3 exclude the median; also called "Tukey's method"). When the
  element is absent PowerPoint treats the method as ``"inclusive"``.
* ``cx:layoutPr/cx:visibility`` carries four boolean flags controlling
  the visual features:

  - ``@meanMarker`` — the "×" / "+" marker at each group's mean.
  - ``@meanLine`` — a line drawn across all group means (rare; off by
    default).
  - ``@nonoutliers`` — every individual non-outlier data point shown
    alongside the box as a small dot (off by default).
  - ``@outliers`` — data points outside 1.5×IQR shown as hollow dots
    (on by default).

  All four default to PowerPoint's internal defaults when the
  ``cx:visibility`` element is absent.

* ``cx:axis`` children are required for box-and-whisker: one
  category-scale axis (``@id="0"`` + ``cx:catScaling``) and one
  value-scale axis (``@id="1"`` + ``cx:valScaling``). Same two-axis
  scaffold as waterfall.
* ``cx:dataId val="0"`` wires the series to the ``cx:data id="0"``
  block — identical to funnel / waterfall.


The ``cx:plotArea`` subset box-and-whisker uses
------------------------------------------------

Mapping from schema to the subset box-and-whisker needs:

* ``cx:plotAreaRegion`` — **required**, one per chart.
* ``cx:plotSurface`` — **required**, empty element; present in every
  PowerPoint-authored box-and-whisker.
* ``cx:series`` — **required (×1)**; exactly one, with
  ``@layoutId="boxWhisker"``.
* ``cx:axis`` — **required (×2)**. One ``cx:catScaling`` category axis
  and one ``cx:valScaling`` value axis — same shape as waterfall.
* ``cx:plotAreaRegion/cx:plotAreaRegionLayout`` — *optional*. Omitted by
  the MVP authoring API.

The two axes can be emitted as bare skeletons for the MVP — no tick
configuration, no title, no text-properties. PowerPoint renders
sensible defaults from the embedded data range.


Series element subset
~~~~~~~~~~~~~~~~~~~~~

Within ``cx:series`` for a box-and-whisker, the MVP needs to emit:

* ``@layoutId`` — **required**, literal ``"boxWhisker"``.
* ``@hidden`` — *optional*, default ``"0"``.
* ``@ownerIdx`` — *optional*, default ``"0"``.
* ``cx:tx/cx:txData/cx:v`` — *optional*. Series name.
* ``cx:layoutPr`` — *optional* overall, but strongly recommended to
  emit an explicit ``cx:statistics`` + ``cx:visibility`` so the chart's
  appearance is deterministic rather than inheriting PowerPoint's
  per-version defaults.
* ``cx:layoutPr/cx:statistics/@quartileMethod`` — *optional*, default
  ``"inclusive"``. The MVP authoring API should accept ``"inclusive"``
  and ``"exclusive"`` as strings (or a small enum).
* ``cx:layoutPr/cx:visibility`` — *optional*, attributes default off /
  on per the table above. The authoring API should expose the four
  flags as booleans.
* ``cx:dataId`` — **required**. ``@val`` matches a ``cx:data/@id``.


``XL_CHART_TYPE.BOX_WHISKER``
-----------------------------

Per F4, ``XL_CHART_TYPE.BOX_WHISKER`` takes the value ``121`` (matching
the Excel VBA ``XlChartType`` constant for box-and-whisker).
``GraphicFrame.chart_type`` reads the first ``cx:series/@layoutId``
and, when it equals ``"boxWhisker"``, returns
``XL_CHART_TYPE.BOX_WHISKER``. This is a post-F4 refinement; today
``chart_type`` returns ``UNSUPPORTED_CHARTEX`` for all chartex kinds and
:attr:`GraphicFrame.chartex_type` is the discriminator.


Authoring API — minimum viable surface
--------------------------------------

Once F4 lands, the public API change for box-and-whisker authoring is
narrow::

  chart_data = CategoryChartData()
  chart_data.categories = ["A", "A", "A", "B", "B", "B"]
  chart_data.add_series("Values", (10, 12, 11, 7, 9, 8))

  chart = slide.shapes.add_chart(
      XL_CHART_TYPE.BOX_WHISKER,
      left, top, width, height,
      chart_data,
  ).chart

The input is a "long" (tall) dataset: each observation is one row,
the category column names the group it belongs to. PowerPoint
groups the rows by category at render time and computes the
min/Q1/median/Q3/max for each group. No precomputation is required.

Optional attributes on :class:`CategoryChartData` (to be added
post-F4):

* ``chart_data.quartile_method = "exclusive"`` — default
  ``"inclusive"``; accepted values ``"inclusive"`` / ``"exclusive"``.
* ``chart_data.show_mean_marker = True``  — default ``True``.
* ``chart_data.show_mean_line = False``   — default ``False``.
* ``chart_data.show_outliers = True``     — default ``True``.
* ``chart_data.show_all_points = False``  — default ``False``.

All five attributes are ignored by non-box-whisker chart kinds (same
pattern as waterfall's ``subtotal_indices``).

Invariants the writer must enforce:

* Exactly one series (``CategoryChartData.add_series`` called once).
* At least one (category, value) pair; PowerPoint silently renders a
  single-observation group as a degenerate "box" with all five
  statistics collapsed.
* Category and value counts match.
* ``quartile_method`` is one of the two legal values.
* Two ``cx:axis`` children — one category-scale (``id="0"``), one
  value-scale (``id="1"``) — always emitted.
* An embedded xlsx is authored alongside and linked via
  ``cx:externalData/@r:id``. Reuses :class:`EmbeddedXlsxPart` from F4.


Out of scope for #1047
----------------------

Explicitly deferred to follow-up issues — not part of the
box-and-whisker MVP:

* **Multiple series** (side-by-side box-and-whisker for comparing
  distributions). PowerPoint supports it via additional ``cx:series``
  children with distinct ``@ownerIdx``; the authoring API stays single-
  series for the MVP.
* **Per-group styling** (``cx:spPr`` / ``cx:dataPoint`` per group). Fill
  / line / outlier-marker overrides. The MVP relies on theme accents.
* **Custom number formats and data labels.** ``cx:dataLabels`` has its
  own extensive subgrammar; box-and-whisker rarely uses them and the
  MVP omits them entirely.
* **Reading the computed quartile statistics** via a public
  ``Plot.categories`` / ``Series.values`` analogue — reading is a
  separate feature.
* **The other chartex kinds not yet analysed in detail** (treemap,
  sunburst, histogram / Pareto, map). Each is its own issue; see
  :doc:`chartex-foundation` for the umbrella list.


Checklist — shipped now vs blocked until F4
-------------------------------------------

**Shipped now (passthrough MVP):**

- [x] :attr:`GraphicFrame.has_chartex` is |True| for a box-and-whisker
      graphic-frame.
- [x] :attr:`GraphicFrame.chartex_type` returns ``"boxWhisker"`` for a
      box-and-whisker graphic-frame.
- [x] :attr:`GraphicFrame.chart_type` returns
      :attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX`.
- [x] The ``cx:chartSpace`` part round-trips verbatim through save /
      reload.
- [x] Regression test
      ``tests/test_issue_1047_box_whisker_passthrough.py`` pins the
      contract.

**Blocked on F4 (structured authoring):**

- [ ] :class:`ChartExPart` registered for
      ``application/vnd.ms-office.chartex+xml`` in
      :data:`pptx.content_type_to_part_class_map`.
- [ ] ``pptx/oxml/chartex/`` subpackage exists with at least the subset
      used by box-and-whisker (``CT_ChartExChartSpace``,
      ``CT_ChartExChartData``, ``CT_ChartExChart``,
      ``CT_ChartExPlotArea``, ``CT_ChartExSeries``,
      ``CT_ChartExExternalData``, ``CT_ChartExAxis``,
      ``CT_ChartExCatScaling``, ``CT_ChartExValScaling``,
      ``CT_ChartExLayoutPr``, ``CT_ChartExStatistics``,
      ``CT_ChartExVisibility``).
- [ ] ``XL_CHART_TYPE.BOX_WHISKER = 121`` defined (replacing the
      ``UNSUPPORTED_CHARTEX`` sentinel for ``layoutId="boxWhisker"``).
- [ ] ``GraphicFrame.chart_type`` returns
      ``XL_CHART_TYPE.BOX_WHISKER`` for a chartex graphic-frame whose
      first series has ``layoutId="boxWhisker"``.
- [ ] ``ExChartPart.new_box_whisker(chart_data, package)`` classmethod
      emits the target XML shape.
- [ ] ``ExChartXmlWriter`` templated for box-and-whisker — builds
      ``cx:chartSpace`` from a :class:`CategoryChartData` input plus
      the five optional attributes above.
- [ ] ``Shapes.add_chart(XL_CHART_TYPE.BOX_WHISKER, ...)`` dispatches to
      ``ExChartPart.new_box_whisker`` rather than ``ChartPart.new``.
- [ ] Embedded xlsx is authored with categories in column A and
      observations in column B, starting at row 2. Reuses
      :class:`EmbeddedXlsxPart` with no modification.
- [ ] Input validation: category count equals value count; at least
      one observation; single series only; ``quartile_method`` is
      ``"inclusive"`` or ``"exclusive"``. Raises :exc:`ValueError`
      otherwise.
- [ ] Round-trip test: author a box-and-whisker via ``add_chart``,
      save, reload, confirm ``chart_type == XL_CHART_TYPE.BOX_WHISKER``
      and the cached categories / values / layoutPr flags survive.
- [ ] HISTORY entry along the lines of:
      *feat: #1047 box-and-whisker chart authoring via*
      ``slide.shapes.add_chart(XL_CHART_TYPE.BOX_WHISKER, ...)``.


Related work
------------

* :doc:`chartex-foundation` — F4 foundation design (the blocker for
  structured authoring).
* :doc:`chartex-funnel` — #305 funnel chart design (the sibling
  chartex-authoring issue; shares most of the writer machinery).
* :doc:`chartex-waterfall` — waterfall design (closest structural
  relative; box-and-whisker inherits waterfall's two-axis scaffold and
  ``cx:layoutPr`` on-series block).
* `#583`_ — chartex parent issue covering all seven kinds.
* `#386`_ — Wave 1 passthrough (chartex parts round-trip as opaque
  bytes). Already in master; the box-and-whisker passthrough MVP
  shipped with this issue layers ``chartex_type`` discrimination on
  top of it.
