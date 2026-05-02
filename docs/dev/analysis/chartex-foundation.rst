
chartex (``cx:``) namespace foundation — F4
===========================================

Foundation item *F4* — the package-level scaffolding required before
python-pptx can read or create Office 2016+ "extended" chart types: funnel,
treemap, sunburst, waterfall, histogram/Pareto, box-and-whisker, and map.
These are the charts introduced with PowerPoint 2016 that live under a new
XML namespace (``cx:``) and a separate part type (``chartEx``) distinct from
the legacy ``c:`` chart grammar python-pptx has targeted since day one.

This note is the design-analysis companion to issue `#583`_ ("New Chart Types
in Office 2016") and to the MVP landed in issue `#386`_ ("chartex
passthrough"). #386 surfaces chartex graphic-frames so they round-trip
without being dropped and exposes the sentinel
:attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX`; F4 is the foundation that will let
a future feature work *read*, *modify*, and *create* these chart types
rather than merely preserve them.

.. _#583: https://github.com/scanny/python-pptx/issues/583
.. _#386: https://github.com/scanny/python-pptx/issues/386


Why a separate foundation
-------------------------

The ``cx:`` (chartex) chart vocabulary is not a superset or a cousin of the
legacy ``c:`` (drawingml/chart) grammar. It is a second chart dialect that
coexists with the original:

* **Namespace**. Legacy charts use
  ``http://schemas.openxmlformats.org/drawingml/2006/chart`` (``c:``).
  Extended charts use ``http://schemas.microsoft.com/office/drawing/2014/chartex``
  (``cx:``). The ``cx:`` namespace is already registered in
  :data:`pptx.oxml.ns._nsmap` (necessary for #386's round-tripping).
* **Part**. Legacy charts live at ``/ppt/charts/chart%d.xml`` with content
  type ``application/vnd.openxmlformats-officedocument.drawingml.chart+xml``.
  Extended charts live at ``/ppt/charts/chartEx%d.xml`` with content type
  ``application/vnd.ms-office.chartex+xml`` (the ``ms-office`` vendor prefix
  betrays its Microsoft-only origin — ECMA-376 never absorbed it).
* **Relationship type**. Legacy charts are related via
  ``.../relationships/chart``. Extended charts use
  ``http://schemas.microsoft.com/office/2014/relationships/chartEx``.
* **Grammar**. The two schemas have genuinely different shapes — different
  root elements (``c:chartSpace`` vs ``cx:chartSpace``), different series
  containers (``c:ser`` under a chart-type element vs ``cx:series`` with a
  ``layoutId`` attribute), different data dimensions (``c:cat`` / ``c:val``
  vs ``cx:strDim`` / ``cx:numDim``). You cannot read a chartex chart with
  the existing :mod:`pptx.chart` classes — the XPaths and element classes
  don't match.

A new chart type cannot be bolted onto the existing ``Chart`` tree; it needs
its own grammar, its own part class, and its own entry point on
``GraphicFrame``. F4 is the minimum set of additions required before
*any* such work becomes possible.


What ships in the MVP already (#386)
------------------------------------

Before F4 proper, issue #386 landed the following to prevent data loss:

* ``cx`` namespace prefix registered in :data:`pptx.oxml.ns._nsmap`.
* :const:`pptx.opc.constants.CONTENT_TYPE.OFC_CHART_EX` constant for the
  ``application/vnd.ms-office.chartex+xml`` content type.
* :const:`pptx.spec.GRAPHIC_DATA_URI_CHARTEX` constant.
* ``GraphicFrame.has_chart`` returns |True| for chartex frames;
  ``GraphicFrame.has_chartex`` distinguishes them; ``GraphicFrame.chart_type``
  returns :attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX` for them; and
  ``GraphicFrame.chart`` raises :exc:`NotImplementedError`. ``shape_type``
  reports :attr:`MSO_SHAPE_TYPE.CHART` in both cases.
* A catch-all fallthrough in ``PartFactory`` (via the base :class:`Part`
  class) preserves the chartex part's bytes on round-trip even though no
  specialised class is registered for the chartex content type.

This is enough to *preserve* chartex charts on load/save but nothing more.
The remaining scope of F4 is the scaffolding required to go beyond that.


What F4 must add
----------------

F4 is a package-and-parse-layer foundation. It does not ship user-visible
chart API — that is left to the feature issues stacked on top of it (the
headline one being #583). Its deliverables are:


1. Specialised ``ChartExPart``
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

A new class in ``pptx/parts/chart.py`` (or a sibling module ``chartex.py``
in the same package) that:

* Declares ``partname_template = "/ppt/charts/chartEx%d.xml"`` — distinct
  from :class:`ChartPart`'s ``/ppt/charts/chart%d.xml``.
* Is registered in ``pptx/__init__.py``'s
  ``content_type_to_part_class_map`` under
  :attr:`CT.OFC_CHART_EX`, so ``PartFactory`` dispatches to it instead of
  the generic :class:`Part`.
* Mirrors ``ChartPart.chart`` / ``ChartPart.chart_workbook`` with
  ``ExChart`` and ``ExChartWorkbook`` analogues. ``ExChartWorkbook`` is
  largely the same code — the embedded xlsx relationship type is still
  ``.../relationships/package`` and the element is ``cx:externalData``
  with an ``r:id`` attribute, so the ``chart_workbook.update_from_xlsx_blob``
  path can be shared via a small mixin or duplicated.
* Provides a ``classmethod new(cls, chart_type, chart_data, package)`` that
  mirrors :meth:`ChartPart.new` but emits a ``cx:chartSpace`` skeleton
  appropriate to the requested chart type. See §4 below for the chart-type
  mapping.


2. OXML element classes (``pptx/oxml/chartex/``)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

A new subpackage, parallel to ``pptx/oxml/chart/``, holding
``xmlchemy``-style descriptors for the chartex elements. The top-level
types, derived from the chartEx XSD (see §3 below) and from PowerPoint
output, are roughly:

``CT_ChartExChartSpace`` (``cx:chartSpace``)
    Root element. Contains one ``cx:chartData`` and one ``cx:chart``; may
    contain ``cx:clrMapOvr``, ``cx:spPr``, ``cx:txPr``, ``cx:externalData``.

``CT_ChartExChartData`` (``cx:chartData``)
    One or more ``cx:data`` blocks. Each ``cx:data`` has a ``@id``
    attribute and any of ``cx:strDim``, ``cx:numDim`` children — the
    chartex analogue of "cache + formula + dimension type".

``CT_ChartExChart`` (``cx:chart``)
    Contains ``cx:title?``, ``cx:plotArea``, ``cx:legend?``,
    ``cx:plotVisOnly?``, ``cx:dispBlanksAs?``.

``CT_ChartExPlotArea`` (``cx:plotArea``)
    Contains ``cx:plotAreaRegion`` (with ``cx:plotSurface`` and one or
    more ``cx:series``), plus zero or more ``cx:axis`` elements.

``CT_ChartExSeries`` (``cx:series``)
    The key element — its ``@layoutId`` attribute is how chartex
    distinguishes chart *kind*. Legal values include ``waterfall``,
    ``funnel``, ``sunburst``, ``treemap``, ``boxWhisker``, ``clusteredColumn``
    (for histogram/Pareto), ``paretoLine`` (Pareto overlay),
    ``regionMap`` (map). Unlike legacy charts, the chart-kind indicator
    is a **series-level** attribute, not a plot-level element — several
    layoutIds can coexist in one chartSpace (e.g. histogram + Pareto line).

``CT_ChartExExternalData`` (``cx:externalData``)
    Carries ``@r:id`` pointing at the embedded xlsx part. Structurally
    identical to ``c:externalData`` — same relationship type, same
    package-part shape — so F4 can reuse :class:`EmbeddedXlsxPart`
    unchanged.

Supporting simple-type attribute classes (``ST_LayoutId``,
``ST_NumericValue``, ``ST_Visibility`` — "hidden"/"shown", etc.) follow the
existing xmlchemy patterns in ``pptx/oxml/chart/``.


3. Spec / schema reference
~~~~~~~~~~~~~~~~~~~~~~~~~~

The chartEx grammar is **not** in the ECMA-376 spec tree the project already
vendors under ``spec/ISO-IEC-29500-*``. It was published by Microsoft
separately as part of the MS-OI29500-2014 addendum / MS-XLSX delta. The
authoritative references are:

* **[MS-XLSX]** appendix, "Extended chart" — Microsoft's published
  protocol documentation for Office 2016 chart extensions. Publicly
  available from https://learn.microsoft.com/openspecs/office_standards/.
* The schema distributed with the Microsoft Office Schemas package as
  ``dml-chartEx.xsd`` (also available via
  ``https://schemas.microsoft.com/office/drawing/2014/chartex``).

Landing F4 should vendor ``dml-chartEx.xsd`` alongside the existing
schemas — proposed location is ``spec/office-2014/xsd/dml-chartEx.xsd``
(new top-level directory for Microsoft-only 2014+ schemas, parallel to
``spec/ISO-IEC-29500-*``). The directory separation keeps it obvious
that chartex is a vendor extension, not an ISO/IEC-blessed part of the
spec.

Until that vendoring lands, the project's chartex work operates from:

* The in-tree fixture ``features/steps/test_files/cht-chartex.pptx``,
  which contains a small waterfall chart authored by PowerPoint (see the
  XML sample in the next section).
* The public Microsoft schema URL above.
* Inspection of the XML written by current PowerPoint clients for each
  of the seven chart kinds.


4. ``XL_CHART_TYPE`` members and a layoutId mapping
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#386 introduced the sentinel ``XL_CHART_TYPE.UNSUPPORTED_CHARTEX`` as a
placeholder for all seven kinds. With F4 in place, the enum grows a
proper member per kind, aligned with the Excel VBA
``XlChartType`` numeric values where Microsoft defined them:

================  =====  =========================================  ===============================
Member            int    ``cx:series/@layoutId``                    Notes
================  =====  =========================================  ===============================
``FUNNEL``        123    ``funnel``
``TREEMAP``       117    ``treemap``
``SUNBURST``      120    ``sunburst``
``WATERFALL``     119    ``waterfall``
``HISTOGRAM``     118    ``clusteredColumn``                        distinguish by axis config
``PARETO``        122    ``paretoLine`` + ``clusteredColumn``       composite (two series)
``BOX_WHISKER``   121    ``boxWhisker``
``REGION_MAP``    140    ``regionMap``                              requires regionLabelLayout
================  =====  =========================================  ===============================

The numeric values match the Excel ``XlChartType`` constants introduced in
Office 2016; they are stable Microsoft-assigned IDs. ``UNSUPPORTED_CHARTEX``
is retained as a fallback sentinel for chartex charts whose series
``layoutId`` is not recognised (or is missing — chartex is still evolving).

``GraphicFrame.chart_type`` gains a new branch that, when
``has_chartex`` is |True|, reads the first ``cx:series/@layoutId`` and maps
it to the corresponding ``XL_CHART_TYPE`` member, falling back to
``UNSUPPORTED_CHARTEX``. Histogram vs Pareto requires a second lookup
(presence of a ``paretoLine`` series → Pareto; otherwise Histogram).


5. Tests and fixtures
~~~~~~~~~~~~~~~~~~~~~

F4 lands with unit tests for:

* ``ChartExPart`` registration and ``.load()`` round-trip (using the existing
  ``cht-chartex.pptx`` fixture).
* Each ``CT_*`` oxml class in isolation, per the project's standard
  ``tests/oxml/test_*.py`` style.
* ``GraphicFrame.chart_type`` mapping for a fixture presentation containing
  one chartex chart of each of the seven kinds (needs a new fixture —
  easiest way to create it is a seven-chart ``.pptx`` authored by PowerPoint
  2016+, not something the library needs to emit from scratch).


What F4 deliberately does not do
--------------------------------

* **No chart-creation API**. Emitting valid ``cx:chartSpace`` XML for each of
  the seven chart types is substantial work and needs its own
  ``ExChartXmlWriter`` family analogous to ``pptx/chart/xmlwriter.py``. F4
  stops at the OXML classes; the writer is feature-level work (#583).
* **No new high-level ``ExChart`` surface**. ``GraphicFrame.chart`` stays
  raising :exc:`NotImplementedError` until there is a real ``ExChart``
  object to return. F4 puts the parsing machinery in place; exposing it
  through a public API is a feature-level decision (likely a new
  :class:`pptx.chart.exchart.ExChart` class).
* **No migration of ``c:chart`` tests**. The existing chart code path stays
  untouched; F4 is additive.


XML specimen: a waterfall chartex
---------------------------------

From the in-tree fixture ``features/steps/test_files/cht-chartex.pptx``,
lightly reformatted:

.. highlight:: xml

Content-type override::

  <Override PartName="/ppt/charts/chartEx1.xml"
            ContentType="application/vnd.ms-office.chartex+xml"/>

Slide relationship::

  <Relationship Id="rId9"
                Type="http://schemas.microsoft.com/office/2014/relationships/chartEx"
                Target="../charts/chartEx1.xml"/>

Slide graphic-frame (inside an ``mc:AlternateContent`` so legacy consumers
see the fallback ``p:sp``)::

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
        <cx:chart r:id="rId9"/>
      </a:graphicData>
    </a:graphic>
  </p:graphicFrame>

The chartEx part itself::

  <cx:chartSpace
      xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <cx:chartData>
      <cx:data id="0">
        <cx:strDim type="cat">
          <cx:f>Sheet1!$A$2:$A$5</cx:f>
          <cx:lvl ptCount="4">
            <cx:pt idx="0">Cat 1</cx:pt>
            <cx:pt idx="1">Cat 2</cx:pt>
            <cx:pt idx="2">Cat 3</cx:pt>
            <cx:pt idx="3">Cat 4</cx:pt>
          </cx:lvl>
        </cx:strDim>
        <cx:numDim type="val">
          <cx:f>Sheet1!$B$2:$B$5</cx:f>
          <cx:lvl ptCount="4" formatCode="General">
            <cx:pt idx="0">10</cx:pt>
            <cx:pt idx="1">20</cx:pt>
            <cx:pt idx="2">-5</cx:pt>
            <cx:pt idx="3">8</cx:pt>
          </cx:lvl>
        </cx:numDim>
      </cx:data>
    </cx:chartData>
    <cx:chart>
      <cx:plotArea>
        <cx:plotAreaRegion>
          <cx:plotSurface/>
          <cx:series layoutId="waterfall" hidden="0" ownerIdx="0">
            <cx:tx><cx:txData><cx:v>Series 1</cx:v></cx:txData></cx:tx>
            <cx:dataLabels pos="outEnd"/>
            <cx:dataId val="0"/>
          </cx:series>
        </cx:plotAreaRegion>
      </cx:plotArea>
    </cx:chart>
  </cx:chartSpace>


Consumers unlocked
------------------

F4 is a building block — it ships no user-visible API on its own. Features
that compose on top of it include:

* `#583`_ — typed creation of funnel, treemap, sunburst, waterfall,
  histogram, box-and-whisker charts via ``slide.shapes.add_chart()``.
* Future histogram binning / Pareto-line configuration API.
* Chart-type read-only access for chartex charts (currently all report
  ``UNSUPPORTED_CHARTEX``).
* `#386`_ can drop its ``NotImplementedError`` and return a real chart
  object when ``has_chartex`` is |True|.


Migration strategy
------------------

F4 is purely additive and does not change the behaviour of existing legacy
chart code. The sentinel ``XL_CHART_TYPE.UNSUPPORTED_CHARTEX`` is retained
across the transition: before F4 it is the only chartex-related value
``chart_type`` can return; after F4 it becomes a fallback for unknown
chartex layoutIds. Callers that pattern-match on
``UNSUPPORTED_CHARTEX`` continue to work; callers that grow
chartex-specific branches pick up the new enum members when they are
ready.
