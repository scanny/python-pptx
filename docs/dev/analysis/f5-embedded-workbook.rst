Foundation F5 - Cross-part embedded-workbook handler
=====================================================

Status
------

**Shipped** on ``feat/foundation-f5-workbook-handler``. F5 layers three
read/write helpers on top of the existing :class:`ChartWorkbook` plumbing
so that chart copy, combo-chart build, and targeted cell edits no longer
need to re-author a fresh ``.xlsx`` workbook from scratch:

* :attr:`pptx.chart.chart.Chart.workbook` — bytes accessor for the
  chart's embedded ``.xlsx``. Reads return ``bytes`` (or ``None`` when the
  chart has no workbook); assignments replace the blob in place,
  preserving the existing ``EmbeddedXlsxPart`` and its rel id.
* :func:`pptx.parts.embeddedpackage.clone_embedded_xlsx` — duplicates a
  chart's embedded workbook into a second chart's package. Each chart
  keeps its own ``EmbeddedXlsxPart`` — PowerPoint expects every chart's
  ``c:externalData/@r:id`` to resolve to a part it alone owns, so sharing
  a single part across charts is not an option.
* :func:`pptx.chart.chart.update_embedded_xlsx_cell` — surgical cell
  write that updates the xlsx blob **and** every ``c:numCache`` /
  ``c:strCache`` entry that references the cell, in a single
  transaction. This avoids the "open the file, click Refresh Data" step
  that would otherwise be required after an external cell edit.

The backing primitives live in :mod:`pptx.chart.xlsx`:

* :class:`WorkbookReader` (pre-existing in Wave 2 #381) — read-only
  per-cell lookups over the existing ``.xlsx`` zip bytes.
* :class:`WorkbookUpdater` (new) — parallel writer that rewrites one
  cell at a time using :mod:`lxml` on the worksheet XML, preserving
  styles, shared strings, docProps, and unrelated cells byte-for-byte.
* :func:`parse_a1_cell`, :func:`_row_col_to_a1` — A1-ref tokeniser
  helpers.

Why F5?
-------

The existing chart-workbook plumbing was always "re-author or replace"
only:

.. code-block:: text

    ChartData -> xlsx_blob (via XlsxWriter)  -> ChartWorkbook.update_from_xlsx_blob
                                             \-> new EmbeddedXlsxPart.blob = ...

There was no way to:

1. **Read** the existing workbook bytes without first discovering the
   private ``chart._workbook.xlsx_part.blob`` path.
2. **Copy** a workbook between charts without re-building the data from
   a :class:`ChartData` proxy (which would drop formulas, comments,
   styles, and any data that isn't surfaced through python-pptx's
   limited :class:`ChartData` API).
3. **Edit** a single cell — callers had to unzip the blob manually, edit
   the worksheet XML by hand, rezip, and push the bytes back, *and* also
   rewrite the matching ``c:numCache``/``c:strCache`` entry so the chart
   didn't render with stale values until PowerPoint itself refreshed.

F5 closes those three gaps with minimal, orthogonal additions.

Primary files
-------------

* ``src/pptx/parts/embeddedpackage.py`` — adds
  :func:`clone_embedded_xlsx`.
* ``src/pptx/chart/xlsx.py`` — adds :func:`parse_a1_cell`,
  :func:`_row_col_to_a1`, :class:`WorkbookUpdater`.
* ``src/pptx/chart/chart.py`` — adds :attr:`Chart.workbook` property,
  :func:`update_embedded_xlsx_cell`, private
  ``_SingleCellCacheRefresher`` helper.
* ``src/pptx/chart/data.py`` — unchanged.  (Listed in the gaps-plan
  primary-files list but F5 ultimately needed nothing there: the
  :class:`ChartData.xlsx_blob` path stays the authoring entry-point,
  while F5 only deals with already-built blobs.)

Design decisions
----------------

1. **Bytes, not file handles.**  :attr:`Chart.workbook` returns
   ``bytes`` so callers can do whatever they want — write to disk, feed
   to ``openpyxl``, diff, round-trip through ``zipfile`` themselves.
   This matches :attr:`ImagePart.blob` / :attr:`Part.blob` precedent in
   the codebase.

2. **Setter delegates to** :meth:`ChartWorkbook.update_from_xlsx_blob`.
   This preserves the issue-#490 "broken rel recovery" behaviour for
   free: a chart with a stale ``c:externalData/@r:id`` that still
   references a now-missing part gets a brand-new rel on first write.

3. **Clone copies bytes + creates a new part.**
   :func:`clone_embedded_xlsx` explicitly does **not** share
   :class:`EmbeddedXlsxPart` instances between charts. Sharing would be
   tempting because ``PACKAGE``-typed rels *can* reference the same
   underlying part from multiple chart parts — but PowerPoint's "Edit
   Data" dialog opens the referenced file for exclusive use, mutates it,
   and writes back.  If two charts point at the same file, editing one
   silently corrupts the other.

4. **Single-cell edits are a single transaction.**
   :func:`update_embedded_xlsx_cell` rewrites the blob *before* it
   rewrites the caches, so a failure in the blob path (e.g. the chart
   has no embedded workbook — ``ValueError``) short-circuits before the
   chart XML is touched. After both sides succeed, :meth:`Chart.workbook`
   and every ``c:numCache``/``c:strCache`` targeting the edited cell
   agree.

5. **Ranges in caches keep their other points.**  The naive "blow away
   the cache, re-populate from the workbook" approach (what
   :meth:`Chart.update_cached_values` does) would be surprising for a
   one-cell edit on a multi-cell range — it would replace *all* the
   cached points, not just the one the caller wrote. The
   ``_SingleCellCacheRefresher`` instead updates the single ``c:pt``
   whose ``@idx`` lines up with the target cell's position in the
   range's ``parse_sheet_range_ref`` cell list, leaving siblings alone.

6. **Inline strings, not shared-strings table.**  When the updater
   writes a string cell, it emits ``t="inlineStr"`` with a nested
   ``<is><t>…</t></is>``.  This avoids having to keep the workbook-wide
   ``xl/sharedStrings.xml`` index in sync on every write, which would
   require either re-indexing every sibling cell that references the
   shared-strings table or appending new entries and leaking stale ones.
   XlsxWriter happily reads ``inlineStr`` cells, and the existing
   :class:`WorkbookReader` already supports them.

7. **Sheet-name fallback to first sheet.**  PowerPoint and XlsxWriter
   both put chart data on a single sheet named ``"Sheet1"``. Files
   authored elsewhere may rename the tab.  Both
   :class:`WorkbookReader` and :class:`WorkbookUpdater` fall back to
   the first sheet when the requested name isn't found, matching the
   existing :meth:`Chart.update_cached_values` semantics.

Schema details
--------------

SpreadsheetML cell element::

    <c r="B2" s="1"><v>1.5</v></c>           <!-- numeric -->
    <c r="B2" t="inlineStr"><is><t>A</t></is></c>  <!-- inline string -->
    <c r="B2" t="s"><v>0</v></c>             <!-- shared-string ref -->
    <c r="B2" t="b"><v>1</v></c>             <!-- boolean -->

The updater always emits the first two forms — numeric for
``int``/``float`` values and inline-string for ``str`` values.

Chart cache/formula shape::

    <c:numRef>
      <c:f>Sheet1!$B$2:$B$4</c:f>
      <c:numCache>
        <c:formatCode>General</c:formatCode>
        <c:ptCount val="3"/>
        <c:pt idx="0"><c:v>1</c:v></c:pt>
        <c:pt idx="1"><c:v>2</c:v></c:pt>
        <c:pt idx="2"><c:v>3</c:v></c:pt>
      </c:numCache>
    </c:numRef>

The range ``Sheet1!$B$2:$B$4`` enumerates cells
``[(2,2), (3,2), (4,2)]`` (row-major). Writing cell ``B3`` finds index
``1`` in that list and rewrites the ``<c:pt idx="1">`` child alone.

Downstream readiness
--------------------

#338 — Combo charts (Wave 3, shipped with inline-literal workaround)
    The MVP writes ``c:numLit``/``c:strLit`` inline so the combo plot
    doesn't need workbook entries. A follow-on change can now use
    :func:`update_embedded_xlsx_cell` to populate cells for the new
    series, then rewrite the combo plot's ``c:ser`` to reference those
    cells — F5 provides the "write a cell, keep caches coherent" half
    of that refactor. The other half (workbook-append: adding columns
    without mangling existing layout) remains future work — F5 is
    deliberately *cell-granularity*, not *column-append*.

#877 — Cross-slide chart copy
    Each copied chart needs its own embedded workbook because
    ``c:externalData`` rels are owned by the chart part.
    :func:`clone_embedded_xlsx` is the primitive the eventual copy-chart
    flow will call: read the source chart's bytes, attach a fresh part
    to the target's package.

#239 — Replace chart data preserving formulas
    :attr:`Chart.workbook` gives callers the raw bytes they can feed to
    an Excel-aware library (openpyxl, xlwings, or the more constrained
    :class:`WorkbookReader`/:class:`WorkbookUpdater` pair); they edit,
    then assign the new bytes back. The setter preserves the existing
    part identity so relationships don't dangle.

#833 — Chart.replace_data with embedded-cell overlay
    The eventual implementation can use
    :func:`update_embedded_xlsx_cell` as the per-cell write primitive,
    and :attr:`Chart.workbook` to seed the workbook from a template
    blob.

Public API
----------

::

    chart.workbook                              # bytes | None (read)
    chart.workbook = xlsx_blob                  # write

    clone_embedded_xlsx(source_part, target_part)   # -> EmbeddedXlsxPart | None

    update_embedded_xlsx_cell(chart, sheet, a1_ref, value)   # -> bytes

``value`` may be ``None`` (clear), ``bool``, ``int``, ``float``, or
``str``. ``a1_ref`` is a single-cell A1 reference without a sheet
qualifier (e.g. ``"B2"`` or ``"$B$2"``, **not** ``"Sheet1!B2"``).

What F5 intentionally does **not** do
-------------------------------------

* **Column-append.**  Adding a new series column (what #338 needs for
  real embedded-cell combo charts) requires re-flowing formatting, row
  dimensions, and shared-string indices across the entire worksheet.
  That's a separate workbook-append primitive — out of scope for F5,
  which sticks to cell-granularity edits.
* **Formula evaluation.**  The updater writes literal values.  Chart
  workbooks authored by PowerPoint never contain formulas (XlsxWriter
  doesn't emit them), and chart caches track raw values anyway, so
  recalculating formulas after a cell write would only matter to
  non-chart consumers of the blob.
* **Shared-strings table maintenance.**  The updater writes inline
  strings so the shared-strings table can stay untouched.  Callers that
  need shared-strings-normalised output should round-trip through a
  full XlsxWriter re-author via :class:`ChartData.xlsx_blob`.
* **Workbook creation from nothing.**  A chart with no ``c:externalData``
  won't grow one from :func:`update_embedded_xlsx_cell` — the caller
  must first assign a starter blob via :attr:`Chart.workbook = ...` (or
  via :meth:`ChartWorkbook.update_from_xlsx_blob`).
