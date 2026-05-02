
Funnel chart (chartex) — design analysis
========================================

Design companion to issue `#305`_ ("Add Funnel Charts"). A *funnel chart*
is one of the seven Office 2016+ "extended" chart kinds that live under the
``cx:`` namespace rather than the legacy ``c:`` grammar. This note covers
the specifics of *authoring* (creating) a funnel chart — the XML shape
python-pptx must be able to emit, and the narrow slice of the chartex
grammar that funnel actually exercises.

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

What is still missing, and what #305 therefore needs, is a specialised
``ChartExPart`` plus the ``cx:`` oxml element tree and a writer family.
None of that has landed. Until it does, the library cannot create a
funnel chart — only preserve one round-tripped from PowerPoint.

The rest of this note assumes F4 will ship substantially as described in
:doc:`chartex-foundation` and focuses on the increment funnel adds on top
of it.


Scope of this note
------------------

The F4 design note (:doc:`chartex-foundation`) covers the *shared*
infrastructure — part class, namespace, content type, the oxml subpackage
layout, the ``XL_CHART_TYPE.FUNNEL = 123`` enum member, and the
``cx:series/@layoutId`` dispatch. This note is narrower: it describes
only the subset of the chartex grammar a funnel chart actually uses, so
the #305 feature work has a concrete target.

Specifically:

* the ``cx:chartSpace`` / ``cx:chart`` / ``cx:plotArea`` shape a
  funnel chart requires,
* the ``cx:series`` shape with ``layoutId="funnel"``,
* the cached-data structure (``cx:chartData``) funnel charts need,
* the minimum viable authoring API on ``Shapes.add_chart(...)``.


Why funnel is the simplest chartex kind
---------------------------------------

Of the seven chartex kinds, funnel is the easiest first target for full
authoring:

* **One series per chart.** Unlike Pareto (histogram + Pareto-line
  overlay) or box-and-whisker (implicit statistical series), a funnel
  chart has exactly one ``cx:series`` element. No composition, no
  multi-layoutId coexistence.
* **One categorical + one numeric dimension.** Funnel uses
  ``cx:strDim type="cat"`` + ``cx:numDim type="val"`` — the simplest
  possible ``cx:chartData`` shape, identical to the waterfall specimen
  already in ``features/steps/test_files/cht-chartex.pptx``.
* **No axes.** ``cx:plotArea`` for a funnel contains no ``cx:axis``
  children. PowerPoint derives the category labels from the series
  itself. Histogram, Pareto, box-and-whisker and waterfall all require
  ``cx:axis`` elements; funnel, treemap and sunburst do not.
* **No regionLabelLayout / region topology data.** Unlike the map chart
  (``regionMap``), funnel needs no supplementary parts.
* **No dataPoint subelements required.** A minimal funnel does not need
  ``cx:dataPoint`` overrides; the default fill/line/series-level
  ``cx:spPr`` is sufficient.

In other words, a funnel is the "hello world" of chartex authoring: the
smallest non-trivial chart the library can learn to emit once F4 ships.


Target XML shape
----------------

Below is the canonical shape PowerPoint emits for a minimal funnel chart
with four categories. The structure is the target for the feature-level
``ExChartXmlWriter`` for funnel.

.. highlight:: xml

Package parts
~~~~~~~~~~~~~

Content type override (authored into ``[Content_Types].xml``)::

  <Override PartName="/ppt/charts/chartEx1.xml"
            ContentType="application/vnd.ms-office.chartex+xml"/>

Slide relationship (``/ppt/slides/_rels/slide1.xml.rels``)::

  <Relationship Id="rId2"
                Type="http://schemas.microsoft.com/office/2014/relationships/chartEx"
                Target="../charts/chartEx1.xml"/>

Slide graphic-frame (wrapped in ``mc:AlternateContent`` so legacy
viewers see a placeholder shape — same pattern PowerPoint uses)::

  <p:graphicFrame>
    <p:nvGraphicFramePr>
      <p:cNvPr id="2" name="Chart 1"/>
      <p:cNvGraphicFramePr/>
      <p:nvPr/>
    </p:nvGraphicFramePr>
    <p:xfrm>
      <a:off x="914400" y="914400"/>
      <a:ext cx="5486400" cy="3657600"/>
    </p:xfrm>
    <a:graphic>
      <a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">
        <cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  r:id="rId2"/>
      </a:graphicData>
    </a:graphic>
  </p:graphicFrame>

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
            <cx:pt idx="0">Awareness</cx:pt>
            <cx:pt idx="1">Interest</cx:pt>
            <cx:pt idx="2">Decision</cx:pt>
            <cx:pt idx="3">Action</cx:pt>
          </cx:lvl>
        </cx:strDim>
        <cx:numDim type="val">
          <cx:f>Sheet1!$B$2:$B$5</cx:f>
          <cx:lvl ptCount="4" formatCode="General">
            <cx:pt idx="0">1000</cx:pt>
            <cx:pt idx="1">600</cx:pt>
            <cx:pt idx="2">250</cx:pt>
            <cx:pt idx="3">80</cx:pt>
          </cx:lvl>
        </cx:numDim>
      </cx:data>
    </cx:chartData>
    <cx:chart>
      <cx:plotArea>
        <cx:plotAreaRegion>
          <cx:plotSurface/>
          <cx:series layoutId="funnel" hidden="0" ownerIdx="0">
            <cx:tx><cx:txData><cx:v>Conversions</cx:v></cx:txData></cx:tx>
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

* ``cx:series/@layoutId="funnel"`` is the *only* element that
  distinguishes this from the waterfall specimen in the F4 note. The
  whole plot-area / data machinery is identical between the two kinds.
  This is a structural confirmation that a funnel writer is a
  one-attribute override on whatever writer a future waterfall
  implementation would use.
* ``cx:dataLabels`` is commonly present on funnels; position
  ``"ctr"`` (centre) is PowerPoint's default. ``"outEnd"`` and
  ``"inBase"`` also appear. This is a candidate for an early option on
  the authoring API.
* ``cx:dataId val="0"`` wires the series to the ``cx:data id="0"``
  block under ``cx:chartData``. Funnel uses a single data block; the
  numeric ``@id`` values are effectively 0-indexed sequence numbers in
  PowerPoint-authored files.
* ``cx:externalData`` points at the embedded ``xl/worksheets/sheet1.xml``
  packaged under ``/ppt/embeddings/Microsoft_Excel_Worksheet%d.xlsx`` —
  identical machinery to legacy charts' ``c:externalData``. The F4 note
  confirms :class:`EmbeddedXlsxPart` can be reused unchanged.


The ``cx:plotArea`` subset funnel uses
--------------------------------------

Mapping from schema to the subset funnel needs:

* ``cx:plotAreaRegion`` — **required**, one per chart.
* ``cx:plotSurface`` — **required**, empty element; present in every
  PowerPoint-authored funnel.
* ``cx:series`` — **required (×1)**; exactly one, with
  ``@layoutId="funnel"``.
* ``cx:axis`` — **not used** for funnel. Funnel has no axes; the
  category labels are derived from ``cx:strDim`` via the series'
  ``cx:dataId`` reference.
* ``cx:plotAreaRegion/cx:plotAreaRegionLayout`` — *optional*. Only
  emitted when the author manually positions the plot-area; the MVP
  authoring API can omit it.

Series element subset
~~~~~~~~~~~~~~~~~~~~~

Within ``cx:series`` for a funnel, the MVP needs to emit:

* ``@layoutId`` — **required**, literal ``"funnel"``.
* ``@hidden`` — *optional*, default ``"0"``. PowerPoint writes it;
  omitting is also valid.
* ``@ownerIdx`` — *optional*, default ``"0"``. Series index; ``"0"``
  for the sole series.
* ``cx:tx/cx:txData/cx:v`` — *optional*. Series name (e.g.
  "Conversions"); if omitted, PowerPoint shows "Series 1".
* ``cx:dataLabels`` — *optional*. Default ``@pos="ctr"`` for funnel.
* ``cx:dataId`` — **required**. ``@val`` matches a ``cx:data/@id``.
* ``cx:spPr`` / ``cx:txPr`` — *optional*. Appearance overrides;
  omitting falls back to theme.
* ``cx:dataPoint`` — **not emitted for MVP**. Per-category overrides;
  not needed for a default funnel.


Cached-data structure for funnel
--------------------------------

Funnel uses a single ``cx:data`` block with exactly two child dimensions:

* **``cx:strDim type="cat"``** — category labels (stage names). At least
  one ``cx:lvl`` with ``@ptCount`` matching the number of categories.
* **``cx:numDim type="val"``** — numeric stage values. Same ``@ptCount``.

Category count mismatches between the two dimensions are rejected by
PowerPoint, so the authoring API needs to validate that
``len(categories) == len(values)`` before emitting.

Formula cells
~~~~~~~~~~~~~

``cx:f`` inside each dimension is a **cached formula** pointing into the
embedded xlsx. For a chart authored from
:class:`~pptx.chart.data.CategoryChartData` the formulas are
deterministic — ``Sheet1!$A$2:$A$N`` and ``Sheet1!$B$2:$B$N`` where
*N* = ``1 + len(values)``. This mirrors :class:`CategoryChartData`'s
existing behaviour for legacy ``c:chart`` parts.

Level elements
~~~~~~~~~~~~~~

``cx:lvl`` is a single-level wrapper in the funnel case; chartex's
multi-level grouping (used by sunburst and treemap) is **not** used for
funnel. This simplifies the authoring API: a funnel is just
(categories, values), no hierarchy.

Number format
~~~~~~~~~~~~~

``cx:lvl/@formatCode`` defaults to ``"General"`` in PowerPoint output.
The authoring API should accept a ``number_format`` keyword mirroring
the legacy chart API.


``XL_CHART_TYPE.FUNNEL``
------------------------

Per F4, ``XL_CHART_TYPE.FUNNEL`` takes the value ``123`` (matching the
Excel VBA ``XlChartType`` constant for funnel). ``GraphicFrame.chart_type``
reads the first ``cx:series/@layoutId`` and, when it equals
``"funnel"``, returns ``XL_CHART_TYPE.FUNNEL``.


Authoring API — minimum viable surface
--------------------------------------

Once F4 lands, the public API change this issue needs is narrow::

  chart = slide.shapes.add_chart(
      XL_CHART_TYPE.FUNNEL,
      left, top, width, height,
      chart_data,          # pptx.chart.data.CategoryChartData
  ).chart

``chart_data`` is the existing ``CategoryChartData`` — one category axis,
one series. Internally the dispatch on ``chart_type`` selects the new
``ExChartPart.new_funnel(...)`` classmethod rather than the legacy
``ChartPart.new(...)``.

What the returned ``chart`` exposes is open — a minimal read-only
``ExChart`` proxy sufficient to satisfy the add-chart contract is enough
for the MVP; the full surface can follow incrementally.

Invariants the writer must enforce:

* Exactly one series (``CategoryChartData.add_series`` called once).
* At least one category / value pair.
* Category and value counts match.
* ``cx:series/@layoutId`` is the literal ``"funnel"``.
* An embedded xlsx is authored alongside and linked via
  ``cx:externalData/@r:id``. Reuses :class:`EmbeddedXlsxPart` from F4.


Out of scope for #305
---------------------

Explicitly deferred to follow-up issues — not part of the funnel MVP:

* **Per-point fill/line overrides** (``cx:dataPoint``) — the legacy
  ``Point.format`` equivalent for chartex. Funnel charts commonly have
  each stage shaded differently; styling is follow-up work.
* **Custom data label number formats and positions beyond the default**
  — ``cx:dataLabels`` has its own extensive subgrammar.
* **Reading a funnel's cached category/value arrays** via a public
  ``Plot.categories`` / ``Series.values`` analogue. Reading is a
  separate feature.
* **Bulk chart copy / clone operations** for chartex parts.
* **The other six chartex kinds.** Each is its own issue (treemap
  (#?), sunburst (#?), waterfall (#?), histogram/Pareto (#?),
  box-and-whisker (#?), map (#?)). Funnel is the recommended first
  kind precisely because it is the simplest.


Checklist — blocked until F4 lands
----------------------------------

This issue is blocked. When F4 (per :doc:`chartex-foundation`) is fully
implemented in master, the following checklist is the concrete scope
for #305:

**F4 prerequisites (verify before unblocking):**

- [ ] :class:`ChartExPart` registered for
      ``application/vnd.ms-office.chartex+xml`` in
      :data:`pptx.content_type_to_part_class_map`.
- [ ] ``pptx/oxml/chartex/`` subpackage exists with ``CT_ChartExChartSpace``,
      ``CT_ChartExChartData``, ``CT_ChartExChart``, ``CT_ChartExPlotArea``,
      ``CT_ChartExSeries``, ``CT_ChartExExternalData`` — or at least the
      subset used by funnel.
- [ ] ``XL_CHART_TYPE.FUNNEL = 123`` defined (replacing the
      ``UNSUPPORTED_CHARTEX`` sentinel for ``layoutId="funnel"``).
- [ ] ``GraphicFrame.chart_type`` returns ``XL_CHART_TYPE.FUNNEL`` for a
      chartex graphic-frame whose first series has
      ``layoutId="funnel"``.

**#305-specific deliverables:**

- [ ] ``ExChartPart.new_funnel(chart_data, package)`` classmethod emits
      the target XML shape (see "Target XML shape" above).
- [ ] ``ExChartXmlWriter`` (or equivalent) templated for funnel — builds
      ``cx:chartSpace`` from a :class:`CategoryChartData` input.
- [ ] ``Shapes.add_chart(XL_CHART_TYPE.FUNNEL, ...)`` dispatches to
      ``ExChartPart.new_funnel`` rather than ``ChartPart.new``.
- [ ] Embedded xlsx is authored with stage names in column A and numeric
      values in column B, starting at row 2. Reuses
      :class:`EmbeddedXlsxPart` with no modification.
- [ ] ``cx:externalData/@r:id`` relationship wired to the embedded xlsx
      part.
- [ ] Input validation: category count equals value count; at least one
      category; single series only. Raises :exc:`ValueError` otherwise.
- [ ] Round-trip test: author a funnel via ``add_chart``, save, reload,
      confirm ``chart_type == XL_CHART_TYPE.FUNNEL`` and the cached
      categories/values survive.
- [ ] Acceptance test (``features/chart-add.feature`` or similar) that
      opens the saved .pptx in an assertion against the pattern used in
      the existing ``cht-chartex.pptx`` fixture.
- [ ] Minimal ``ExChart`` / ``ExSeries`` read-proxy sufficient for
      ``add_chart`` to return a non-opaque object. Can be read-only for
      the MVP.
- [ ] HISTORY entry along the lines of:
      *feat: #305 funnel chart authoring via*
      ``slide.shapes.add_chart(XL_CHART_TYPE.FUNNEL, ...)``.

**Deliberately deferred (not part of #305):**

- Per-stage fill overrides (``cx:dataPoint``).
- Non-default data-label positions / formats.
- Read-side introspection of the cached categories/values.


Test-fixture plan
-----------------

Two fixtures are needed:

* **Authored-by-PowerPoint funnel fixture** — a .pptx containing a
  minimal four-stage funnel chart, saved by PowerPoint 2019+ (2016 will
  do — the format hasn't changed). Used for round-trip parsing tests
  and as the golden reference for the XML shape the writer must
  produce. Proposed location:
  ``features/steps/test_files/cht-chartex-funnel.pptx``.
* **Hand-crafted minimal XML fixture** — an XML string in
  ``tests/chart/fixtures/chartex_funnel_basic.xml`` (new directory
  paralleling ``tests/chart/fixtures/``) containing just the
  ``cx:chartSpace`` emitted by ``ExChartPart.new_funnel(...)`` for a
  canonical input. Golden-file tested against the writer output.


Related work
------------

* :doc:`chartex-foundation` — F4 foundation design (the blocker for this
  issue).
* `#583`_ — chartex parent issue covering all seven kinds.
* `#386`_ — Wave 1 passthrough (chartex parts round-trip as opaque
  bytes). Already in master.
