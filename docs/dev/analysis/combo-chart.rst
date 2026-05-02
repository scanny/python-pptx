Combo chart (multi-plot chart) — design analysis
================================================

Status
------

**MVP shipped (``feat/issue-338-combo-charts``).** The MVP adds a new
:meth:`pptx.chart.chart.Chart.add_plot` method that appends an additional
``c:{x}Chart`` element inside the chart's existing ``c:plotArea``,
reusing the existing axes. Series values for the added plot are emitted
as *inline literals* (``c:numLit`` for values, ``c:strLit`` for category
labels) rather than as ``c:numRef`` / ``c:strRef`` references into the
embedded Excel workbook. The series name uses ``c:tx/c:strRef/c:strCache``
with a deliberately-unresolvable ``c:f`` so existing readers (including
:class:`pptx.chart.series.Series.name`) keep working via the cache.

Supported chart types for the added plot (MVP):

* ``BAR_CLUSTERED``, ``COLUMN_CLUSTERED``
* ``BAR_STACKED``, ``COLUMN_STACKED``
* ``BAR_STACKED_100``, ``COLUMN_STACKED_100``
* ``LINE``, ``LINE_MARKERS``
* ``LINE_STACKED``, ``LINE_STACKED_100``
* ``LINE_MARKERS_STACKED``, ``LINE_MARKERS_STACKED_100``

Other chart types raise :class:`NotImplementedError`. The first (pre-existing)
plot is unaffected.

Why inline literals? (F5 dependency)
------------------------------------

The natural implementation would:

1. Re-author the chart's embedded ``.xlsx`` workbook so it carries one
   extra column per new series (plus any shared category column), and
2. Emit ``c:numRef`` / ``c:strRef`` elements pointing at those new
   cells in the re-authored workbook.

Step 1 requires a **workbook-append** primitive — read the existing
``.xlsx`` blob, merge new column(s) preserving the existing sheet layout,
re-emit the blob. The existing :class:`pptx.chart.xlsx.CategoryWorkbookWriter`
is a *write-from-scratch* helper: it lays out a fresh sheet from a
:class:`pptx.chart.data.CategoryChartData` and has no concept of an
input workbook to append into. Adding that primitive — bridging
``xlsxwriter`` (write-only, no reader) with something that can read an
existing sheet, figure out the next free column, and write new data
while preserving all formatting, styles, shared-strings, and name
manager entries — is the scope of the **F5 (embedded-workbook handler)
foundation** tracked alongside this issue.

Until F5 lands, synchronizing the workbook on every
:meth:`Chart.add_plot` call would either:

* Re-serialize the entire chart data (existing plots + new plot) through
  ``CategoryWorkbookWriter``, overwriting the original workbook. This
  loses any formatting, macros, or external data connections the
  authored workbook may have carried — a silent regression for
  round-trip users.
* Drop straight into a hand-rolled ``openpyxl``-based patcher, which
  adds a new heavy dependency and duplicates what F5 is planned to
  provide.

Using inline literals (``c:numLit``/``c:strLit``) side-steps the issue
entirely: the schema (``CT_NumDataSource`` / ``CT_AxDataSource`` in
``dml-chart.xsd``) defines ``numRef`` and ``numLit`` as alternatives in
an ``xsd:choice`` — both are first-class. PowerPoint, LibreOffice, and
Keynote render inline-literal series correctly; the only user-visible
difference is that PowerPoint's *Edit Data* dialog shows only the
original plot's data range (the added plot's values are in the chart
XML, not the workbook). This is a tolerable trade-off for an MVP: the
chart is correct, the file round-trips cleanly, and nobody sees stale
numbers.

F5 follow-up: what the foundation needs to expose
-------------------------------------------------

Once F5 lands, the MVP should be upgraded to write the added plot's
series into the embedded workbook and swap ``c:numLit`` / ``c:strLit``
for ``c:numRef`` / ``c:strRef``. The minimal API F5 needs to expose for
this is:

``ChartWorkbook.next_free_column(sheet_name="Sheet1") -> int``
    Return the 1-based index of the first column on *sheet_name* that is
    not used by any existing chart series or category. Needed so the
    added plot's cell ranges don't collide with the original plot's.

``ChartWorkbook.append_series(sheet_name, col_idx, header, values, number_format=None)``
    Append a column to *sheet_name* at *col_idx* with *header* in row 1
    and numeric *values* in rows 2..N+1. Should preserve the existing
    workbook's shared strings, formatting, and other worksheets.
    Returns the Excel-style range ref (e.g. ``"Sheet1!$D$2:$D$5"``)
    suitable for dropping into a ``c:f`` element.

``ChartWorkbook.categories_ref(sheet_name="Sheet1") -> str | None``
    Return the Excel range reference for the existing category column
    (e.g. ``"Sheet1!$A$2:$A$5"``), or |None| if the sheet has no
    category column yet. The added plot can point its ``c:cat/c:strRef``
    at the same range as the original plot rather than duplicating the
    labels.

``ChartWorkbook.series_refs() -> list[dict]`` *(optional but useful)*
    Enumerate existing series as ``{"name_ref": ..., "values_ref": ...,
    "cat_ref": ...}`` so callers can inspect the current layout and
    validate that category labels match before appending.

With that, :meth:`Chart.add_plot` would change from "build
``c:numLit``/``c:strLit`` elements" to "call
``workbook.append_series(...)`` for each series in *chart_data*, then
emit ``c:numRef`` elements carrying the returned range refs", mirroring
what :class:`pptx.chart.xmlwriter._CategorySeriesXmlWriter` does for the
initial plot.

Affected code
-------------

* :class:`pptx.chart.xmlwriter._PlotFragmentBuilder` — dispatcher
* :class:`pptx.chart.xmlwriter._BarPlotFragmentBuilder` — ``c:barChart`` emitter
* :class:`pptx.chart.xmlwriter._LinePlotFragmentBuilder` — ``c:lineChart`` emitter
* :meth:`pptx.chart.chart.Chart.add_plot` — public entry point

Schema references
-----------------

* ``CT_NumDataSource`` (dml-chart.xsd lines 62–70): ``numRef`` XOR ``numLit``
* ``CT_AxDataSource`` (dml-chart.xsd lines 120–136): ``strRef`` XOR ``strLit`` XOR ``numRef`` XOR ``numLit`` XOR ``multiLvlStrRef``
* ``CT_SerTx`` (dml-chart.xsd lines 131–138): ``strRef`` XOR bare ``c:v``

Testing
-------

Unit tests: :file:`tests/chart/test_chart.py` (``DescribeChart`` —
``it_can_add_a_line_plot_to_an_existing_bar_chart_338`` and siblings);
:file:`tests/chart/test_xmlwriter.py` (``Describe_PlotFragmentBuilder``).

Behave scenario: :file:`features/cht-combo.feature` —
*Add a line plot on top of a column chart (issue #338)*.

Verification: round-trip through :class:`pptx.Presentation` preserves
both plots' series names, values, and axis IDs. PowerPoint desktop and
LibreOffice Impress open the generated file without a *Repair* prompt.
