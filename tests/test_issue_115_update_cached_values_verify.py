# pyright: reportPrivateUsage=false

"""Regression test for issue #115 — ``Chart.update_cached_values()``.

Issue #115 (https://github.com/scanny/python-pptx/issues/115) asked for a
programmatic way to refresh the values shown by a chart after its backing
data has been changed:

  > even with linked graph, one has to select every single chart and click
  > update once per chart. this is tedious work for weekly presentation
  > with lot of charts. It would be great if its possible to be done in a
  > programatical way. update ppt with data directly from obdc database.

The reporter's workflow is the classic "data warehouse → PowerPoint
deck" pipeline: a template deck carries charts whose embedded ``.xlsx``
is overwritten with fresh numbers from an external data source (ODBC /
another workbook / any out-of-band producer) and the deck is re-saved.
With stock PowerPoint behaviour the re-saved deck still renders the
*old* cached numbers until each chart is individually selected and
*Refresh Data* (``Alt+F5``) is clicked in the ribbon.

Wave 2 #381 shipped :meth:`Chart.update_cached_values`, which re-reads
the embedded workbook, resolves every ``<c:f>`` cell reference under the
chart, and rewrites the sibling ``<c:numCache>`` / ``<c:strCache>``
subtrees so the chart displays the workbook's current values without
PowerPoint interaction. This is exactly the #115 ask — a single Python
call replaces the manual per-chart refresh step.

#115 is therefore verified-and-closed by #381. The scenarios below
exercise the reporter's end-to-end workflow (author a chart, externally
rewrite its embedded workbook, call
:meth:`Chart.update_cached_values`, save, reopen, confirm the rendered
values match the new workbook) plus the edge cases called out in the
API docstring (no-op on charts without an embedded workbook; unresolved
cell references left untouched).
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from pptx import Presentation
from pptx.chart.chart import Chart
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

# -- helpers ----------------------------------------------------------------


def _author_single_series_column_chart(prs, categories, series_name, values):
    """Add a single-series column chart to *prs*'s first slide and return it.

    Mirrors the one-call snippet the #115 reporter's template-deck
    workflow would start from.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = list(categories)
    chart_data.add_series(series_name, tuple(values))
    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        chart_data,
    )
    return graphic_frame.chart


def _externally_rewrite_workbook(chart, *, numeric_factor=None, string_replacements=None):
    """Externally overwrite the embedded xlsx blob on *chart*.

    Simulates the reporter's ODBC-fed update: the embedded workbook is
    replaced with one that has different cell values without going
    through :meth:`Chart.replace_data`. Only ``s="1"``-styled numeric
    cells (the style XlsxWriter emits for chart data numbers) and shared
    strings are touched so that workbook structure is preserved.

    Returns the new blob for caller-side inspection.
    """
    blob = chart._workbook.xlsx_part.blob
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        files = {name: zf.read(name) for name in zf.namelist()}

    if numeric_factor is not None:
        sheet = files["xl/worksheets/sheet1.xml"].decode()

        def _scale(m):
            addr = m.group("addr")
            try:
                new_val = float(m.group("val")) * numeric_factor
            except ValueError:
                return m.group(0)
            return '<c r="%s" s="1"><v>%s</v></c>' % (addr, new_val)

        sheet = re.sub(
            r'<c r="(?P<addr>[A-Z]+\d+)" s="1"><v>(?P<val>[^<]+)</v></c>',
            _scale,
            sheet,
        )
        files["xl/worksheets/sheet1.xml"] = sheet.encode()

    if string_replacements:
        ss = files["xl/sharedStrings.xml"].decode()
        for old, new in string_replacements.items():
            ss = ss.replace("<t>%s</t>" % old, "<t>%s</t>" % new)
        files["xl/sharedStrings.xml"] = ss.encode()

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    new_blob = buf.getvalue()
    chart._workbook.xlsx_part.blob = new_blob
    return new_blob


# -- fixtures --------------------------------------------------------------


@pytest.fixture
def _restore_part_factory():
    """Guard against cross-test pollution of `PartFactory.part_type_for`.

    ``tests/opc/test_package.py::DescribePartFactory`` overwrites entries
    in the global ``PartFactory.part_type_for`` map during unit testing
    and does not restore them. Our round-trip scenarios rely on the
    real ``ChartPart`` registration so re-open after save yields a
    :class:`Chart` rather than a generic ``XmlPart`` mock. This fixture
    mirrors the same guard used by ``test_issue_303_multi_axes_verify``.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.parts.embeddedpackage import EmbeddedXlsxPart
    from pptx.parts.image import ImagePart
    from pptx.parts.media import MediaPart
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.slide import (
        NotesMasterPart,
        NotesSlidePart,
        SlideLayoutPart,
        SlideMasterPart,
        SlidePart,
    )

    saved = dict(PartFactory.part_type_for)
    expected = {
        CT.PML_PRESENTATION_MAIN: PresentationPart,
        CT.PML_PRES_MACRO_MAIN: PresentationPart,
        CT.PML_TEMPLATE_MAIN: PresentationPart,
        CT.PML_SLIDESHOW_MAIN: PresentationPart,
        CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
        CT.PML_NOTES_MASTER: NotesMasterPart,
        CT.PML_NOTES_SLIDE: NotesSlidePart,
        CT.PML_SLIDE: SlidePart,
        CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
        CT.PML_SLIDE_MASTER: SlideMasterPart,
        CT.DML_CHART: ChartPart,
        CT.SML_SHEET: EmbeddedXlsxPart,
        CT.JPEG: ImagePart,
        CT.PNG: ImagePart,
        CT.MP4: MediaPart,
    }
    PartFactory.part_type_for.update(expected)
    try:
        yield
    finally:
        PartFactory.part_type_for.clear()
        PartFactory.part_type_for.update(saved)


# -- the verification suite ------------------------------------------------


class DescribeIssue115UpdateCachedValuesVerify(object):
    """#115 programmatic chart-refresh verify-and-close via #381.

    Pins :meth:`Chart.update_cached_values` against the reporter's
    end-to-end workflow (external workbook rewrite, programmatic cache
    refresh, save, reopen) plus the edge cases the docstring calls
    out.
    """

    # -- reporter's end-to-end workflow -----------------------------------

    def it_refreshes_cached_values_after_external_workbook_rewrite(self):
        # -- the #115 ask verbatim: author a chart whose embedded xlsx
        # -- is rewritten by an external producer, then programmatically
        # -- refresh so the rendered values match the workbook.
        prs = Presentation()
        chart = _author_single_series_column_chart(
            prs,
            categories=("Old-A", "Old-B", "Old-C"),
            series_name="Old-Series",
            values=(1.0, 2.0, 3.0),
        )

        _externally_rewrite_workbook(
            chart,
            numeric_factor=10,
            string_replacements={
                "Old-A": "New-A",
                "Old-B": "New-B",
                "Old-C": "New-C",
                "Old-Series": "New-Series",
            },
        )

        # -- before the refresh the cache still carries the old values --
        series = chart.series[0]
        assert tuple(series.values) == (1.0, 2.0, 3.0)
        assert series.name == "Old-Series"
        assert [c.label for c in chart.plots[0].categories] == [
            "Old-A",
            "Old-B",
            "Old-C",
        ]

        # -- the single API call the reporter wanted --
        chart.update_cached_values()

        # -- cached series values now match the workbook (10×) --
        assert tuple(series.values) == (10.0, 20.0, 30.0)
        # -- cached series name matches the workbook --
        assert series.name == "New-Series"
        # -- cached category labels match the workbook --
        assert [c.label for c in chart.plots[0].categories] == [
            "New-A",
            "New-B",
            "New-C",
        ]

    def it_persists_the_refreshed_values_through_save_and_reopen(self, _restore_part_factory):
        # -- the reporter's full pipeline: refresh, save, reopen. The
        # -- reopened deck must display the new numbers without a manual
        # -- refresh in PowerPoint.
        prs = Presentation()
        chart = _author_single_series_column_chart(
            prs,
            categories=("Q1", "Q2", "Q3", "Q4"),
            series_name="Revenue",
            values=(100.0, 200.0, 300.0, 400.0),
        )

        _externally_rewrite_workbook(chart, numeric_factor=2)
        chart.update_cached_values()

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        chart2 = next(s for s in prs2.slides[0].shapes if s.has_chart).chart

        # -- reopened chart carries the refreshed cached values --
        assert tuple(chart2.series[0].values) == (200.0, 400.0, 600.0, 800.0)

    def it_lets_the_reporter_refresh_many_charts_in_one_pass(self):
        # -- the #115 pain point: "weekly presentation with lot of charts"
        # -- × "one has to select every single chart and click update".
        # -- Pin that a plain Python loop over every chart on every slide
        # -- is sufficient — no per-chart manual step required.
        prs = Presentation()
        charts = []
        for i in range(3):
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            chart_data = CategoryChartData()
            chart_data.categories = ["a", "b"]
            chart_data.add_series("S%d" % i, (1.0 + i, 2.0 + i))
            gf = slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED,
                Inches(1),
                Inches(1),
                Inches(6),
                Inches(4),
                chart_data,
            )
            charts.append(gf.chart)

        # -- each chart's embedded xlsx is rewritten by the external feed --
        for chart in charts:
            _externally_rewrite_workbook(chart, numeric_factor=100)

        # -- the user's "refresh all" loop: one call per chart --
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_chart:
                    shape.chart.update_cached_values()

        # -- every chart now renders the new values --
        assert tuple(charts[0].series[0].values) == (100.0, 200.0)
        assert tuple(charts[1].series[0].values) == (200.0, 300.0)
        assert tuple(charts[2].series[0].values) == (300.0, 400.0)

    # -- docstring-called-out edge cases ----------------------------------

    def it_is_a_noop_on_a_chart_without_an_embedded_workbook(self):
        # -- "This is a no-op for charts that do not have an embedded
        # -- workbook (i.e. <c:externalData> is absent from the chart
        # -- XML)." Pin that calling it on such a chart does not raise
        # -- and does not mutate the chart XML.
        prs = Presentation()
        chart = _author_single_series_column_chart(
            prs,
            categories=("A", "B"),
            series_name="S",
            values=(1.0, 2.0),
        )
        # -- strip the c:externalData element so the chart no longer has
        # -- an embedded xlsx (mirrors a chart pasted from a pre-2007
        # -- .xls, or one whose embedded-package rel was stripped by
        # -- another client — see issue #490). --
        externalData = chart._chartSpace.xpath(".//c:externalData")
        for ed in externalData:
            ed.getparent().remove(ed)
        assert chart.workbook is None  # -- precondition --

        xml_before = chart._chartSpace.xml

        chart.update_cached_values()  # -- must not raise --

        assert chart._chartSpace.xml == xml_before

    def it_leaves_unresolved_cell_references_untouched(self):
        # -- "When a cell reference cannot be resolved […] the
        # -- corresponding cache is left untouched." Pin this by pointing
        # -- one numRef at a sheet that does not exist in the workbook:
        # -- the other (resolvable) cache must still be refreshed.
        prs = Presentation()
        chart = _author_single_series_column_chart(
            prs,
            categories=("A", "B", "C"),
            series_name="S",
            values=(1.0, 2.0, 3.0),
        )

        # -- rewrite the workbook so a refresh _would_ change values --
        _externally_rewrite_workbook(chart, numeric_factor=10)

        # -- then break one numRef by pointing it at a nonexistent sheet.
        # -- We target the cat strRef (categories) here because that gives
        # -- a clean "this reference is unresolvable" signal without
        # -- breaking the val numRef the test also inspects.
        C_NS = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        cat_f = chart._chartSpace.find(".//%scat/%sstrRef/%sf" % (C_NS, C_NS, C_NS))
        assert cat_f is not None, "precondition: cat strRef should exist"
        original_cat_formula = cat_f.text
        cat_f.text = "BogusSheet!$A$2:$A$4"
        # -- the cat strCache carries the original "A","B","C" labels --
        original_cat_cache_xml = cat_f.getparent().find(C_NS + "strCache")
        assert original_cat_cache_xml is not None
        original_cat_labels = [
            pt.findtext(C_NS + "v") for pt in original_cat_cache_xml.findall(C_NS + "pt")
        ]
        assert original_cat_labels == ["A", "B", "C"]

        chart.update_cached_values()

        # -- the val numCache WAS refreshed (resolvable reference) --
        assert tuple(chart.series[0].values) == (10.0, 20.0, 30.0)
        # -- the cat strCache was NOT touched (reference unresolvable) --
        refreshed_cat_labels = [c.label for c in chart.plots[0].categories]
        assert refreshed_cat_labels == ["A", "B", "C"], refreshed_cat_labels
        # -- restore formula so teardown doesn't leak broken state --
        cat_f.text = original_cat_formula

    def it_is_safe_to_call_repeatedly(self):
        # -- idempotency: a second call with no workbook change should
        # -- leave the already-fresh cache alone.
        prs = Presentation()
        chart = _author_single_series_column_chart(
            prs,
            categories=("A", "B"),
            series_name="S",
            values=(1.0, 2.0),
        )
        _externally_rewrite_workbook(chart, numeric_factor=5)

        chart.update_cached_values()
        first_xml = chart._chartSpace.xml

        chart.update_cached_values()
        second_xml = chart._chartSpace.xml

        assert first_xml == second_xml
        assert tuple(chart.series[0].values) == (5.0, 10.0)

    # -- XY (scatter) coverage --------------------------------------------

    def it_refreshes_caches_on_an_xy_scatter_chart(self):
        # -- pin that the method is not category-chart-specific; the
        # -- reporter's "lot of charts" deck likely contains a mix.
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        chart_data = XyChartData()
        s = chart_data.add_series("Scatter")
        s.add_data_point(1.0, 10.0)
        s.add_data_point(2.0, 20.0)
        s.add_data_point(3.0, 30.0)
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.XY_SCATTER,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            chart_data,
        ).chart

        _externally_rewrite_workbook(chart, numeric_factor=2)

        chart.update_cached_values()

        # -- XY series values correspond to the Y column; after ×2 they
        # -- should read (20, 40, 60). --
        assert tuple(chart.series[0].values) == (20.0, 40.0, 60.0)

    # -- public API surface -----------------------------------------------

    def it_exposes_update_cached_values_on_Chart(self):
        # -- belt-and-braces: the exact method name the reporter needed. --
        assert hasattr(Chart, "update_cached_values")
        assert callable(Chart.update_cached_values)
