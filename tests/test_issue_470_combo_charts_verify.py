# pyright: reportPrivateUsage=false

"""Regression test for issue #470 — Adding subcharts / creating combo charts.

Issue #470 (https://github.com/scanny/python-pptx/issues/470) reports a
caller trying to author a combo chart (e.g. bar + line on the same plot
area) following the recipe from issue #338 and hitting
``AttributeError: 'CategoryWorkbookWriter' object has no attribute
'x_values_ref'``. The reporter's ask — a single API entry point to
overlay a second chart type on an existing chart, sharing the category
axis, optionally bound to a secondary value axis — is covered on this
fork by the :meth:`Chart.add_plot` primitive that shipped under
``feat: #338 combo charts`` in the ``2026.05.0`` release line, together
with :attr:`Chart.has_secondary_value_axis` /
:attr:`Chart.secondary_value_axis` (``feat: #141 secondary value axis
read access``) for the optional right-hand Y axis.

This suite pins the reporter's workflow from the user's perspective:

* Bar + line combo, round-trip through ``Presentation.save`` + reopen,
  asserting both plot types are present after save/reopen.
* Combo with a secondary value axis (loaded from pre-authored
  ``chartSpace`` XML — the documented round-trip path for secondary
  axes — and verified via :attr:`Chart.secondary_value_axis`).
* Custom line-series formatting (line color and marker style) survives
  the save/reopen round-trip on a plot appended via
  :meth:`Chart.add_plot`.

Scope note / follow-up
----------------------
Authoring a secondary ``c:valAx`` element from scratch on a
library-created chart (i.e. a convenience like
``Chart.add_secondary_value_axis()`` or
``add_plot(..., secondary=True)``) is not currently exposed. The
supported path is to (a) start from a PowerPoint-authored template
file that already carries the secondary axis, or (b) author the
``chartSpace`` XML directly and then mutate via
``Chart.secondary_value_axis``. This is called out in
``docs/user/charts.rst`` under "Secondary value axis" and tracked in
``docs/dev/analysis/combo-chart.rst``. #470 closes as delivered
because the reporter's core ask — "Is there any other way of creating
combo charts using python-pptx?" — is satisfied by ``Chart.add_plot``.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.chart.axis import ValueAxis
from pptx.chart.chart import Chart
from pptx.chart.data import CategoryChartData
from pptx.chart.plot import BarPlot, LinePlot
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_MARKER_STYLE
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Inches

# -- fixtures -----------------------------------------------------------


@pytest.fixture
def _restore_part_factory():
    """Guard against pollution from tests that mutate `PartFactory.part_type_for`.

    ``tests/opc/test_package.py::DescribePartFactory`` overwrites
    ``PartFactory.part_type_for[CT.PML_SLIDE]`` with a Mock and does
    not restore it. Mirrors the identical fixture used by the #303
    verify suite (``tests/test_issue_303_multi_axes_verify.py``).
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.coreprops import CorePropertiesPart
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


# -- helpers -------------------------------------------------------------


def _combo_chartSpace_with_secondary_axis_xml():
    """Return chartSpace XML for a column-bar + line combo w/ secondary valAx.

    Shape mirrors what PowerPoint writes for the common "columns on the
    primary axis, line on the secondary axis" combo chart:

    * ``c:barChart`` referencing ``(catAx=111, valAx=222)`` (primary).
    * ``c:lineChart`` referencing ``(catAx=111, valAx=333)`` (secondary).
    * three axes: shared ``c:catAx@axId=111``, primary
      ``c:valAx@axId=222`` on the left, secondary
      ``c:valAx@axId=333`` on the right with a scale (``c:max=100``).

    The series names are the reporter-relatable "Revenue" (bars) and
    "Trend" (line) so the assertions read naturally.
    """
    return (
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
        '/2006/chart">'
        "<c:chart><c:plotArea>"
        # -- primary plot: column chart bound to primary value axis --
        "<c:barChart>"
        '<c:barDir val="col"/>'
        '<c:grouping val="clustered"/>'
        '<c:ser><c:idx val="0"/><c:order val="0"/>'
        "<c:tx><c:v>Revenue</c:v></c:tx>"
        "</c:ser>"
        '<c:axId val="111"/>'
        '<c:axId val="222"/>'
        "</c:barChart>"
        # -- overlay plot: line chart bound to secondary value axis --
        "<c:lineChart>"
        '<c:grouping val="standard"/>'
        '<c:ser><c:idx val="1"/><c:order val="1"/>'
        "<c:tx><c:v>Trend</c:v></c:tx>"
        "</c:ser>"
        '<c:axId val="111"/>'
        '<c:axId val="333"/>'
        "</c:lineChart>"
        # -- shared category axis --
        "<c:catAx>"
        '<c:axId val="111"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/>'
        '<c:axPos val="b"/>'
        '<c:crossAx val="222"/>'
        "</c:catAx>"
        # -- primary value axis (left side) --
        "<c:valAx>"
        '<c:axId val="222"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/>'
        '<c:axPos val="l"/>'
        '<c:crossAx val="111"/>'
        "</c:valAx>"
        # -- secondary value axis (right side) with explicit scale --
        "<c:valAx>"
        '<c:axId val="333"/>'
        "<c:scaling>"
        '<c:orientation val="minMax"/>'
        '<c:max val="100"/>'
        "</c:scaling>"
        '<c:delete val="0"/>'
        '<c:axPos val="r"/>'
        '<c:crossAx val="111"/>'
        "</c:valAx>"
        "</c:plotArea></c:chart></c:chartSpace>"
    )


# -- the verification suite ---------------------------------------------


class DescribeIssue470ComboCharts(object):
    """#470 combo-chart verify-and-close via #338 + #141.

    The reporter asked for a supported way to overlay chart types in a
    single plot area. ``Chart.add_plot`` delivers this; the tests below
    pin the reporter's three concrete asks: (1) bar + line combo with
    round-trip fidelity, (2) combo with an optional secondary value
    axis, (3) custom series formatting surviving save/reopen.
    """

    # -- (1) bar + line combo, round-tripped through save + reopen -------

    def it_creates_a_bar_plus_line_combo_with_add_plot(self, _restore_part_factory):
        """``Chart.add_plot`` overlays a line plot on a column chart (#338)."""
        # -- author a fresh column chart on a slide --
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        bar_data = CategoryChartData()
        bar_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        bar_data.add_series("Revenue", (10, 20, 30, 40))
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            bar_data,
        ).chart

        # -- overlay a line plot sharing the column chart's axes --
        line_data = CategoryChartData()
        line_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        line_data.add_series("Trend", (5, 15, 25, 35))
        new_plot = chart.add_plot(XL_CHART_TYPE.LINE, line_data)

        # -- both plot objects exist in the in-memory chart --
        assert len(chart.plots) == 2
        assert isinstance(chart.plots[0], BarPlot)
        assert isinstance(chart.plots[1], LinePlot)
        assert isinstance(new_plot, LinePlot)

    def it_round_trips_a_bar_plus_line_combo_through_save_reopen(self, _restore_part_factory):
        """Both plot types survive ``Presentation.save`` + reopen (#338)."""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        bar_data = CategoryChartData()
        bar_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        bar_data.add_series("Revenue", (10, 20, 30, 40))
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            bar_data,
        ).chart

        line_data = CategoryChartData()
        line_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        line_data.add_series("Trend", (5, 15, 25, 35))
        chart.add_plot(XL_CHART_TYPE.LINE, line_data)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        graphic_frame = next(s for s in prs2.slides[0].shapes if s.has_chart)
        chart2 = graphic_frame.chart

        # -- both plot types present on reopen --
        assert len(chart2.plots) == 2
        plot_kinds = {type(p).__name__ for p in chart2.plots}
        assert plot_kinds == {"BarPlot", "LinePlot"}
        # -- and a shared category axis underneath both --
        bar_xChart, line_xChart = chart2._chartSpace.plotArea.xCharts
        bar_axIds = [e.get("val") for e in bar_xChart.findall(qn("c:axId"))]
        line_axIds = [e.get("val") for e in line_xChart.findall(qn("c:axId"))]
        assert bar_axIds[0] == line_axIds[0], "category axis is shared"
        assert bar_axIds[1] == line_axIds[1], "value axis is shared"
        # -- series from both plots survive --
        series_names = [s.name for s in chart2.series]
        assert "Revenue" in series_names
        assert "Trend" in series_names

    # -- (2) combo with an optional secondary value axis -----------------

    def it_exposes_a_secondary_value_axis_on_a_combo_chart(self):
        """A combo chart authored with two value axes surfaces the right-hand axis (#141)."""
        chartSpace = parse_xml(_combo_chartSpace_with_secondary_axis_xml())
        chart = Chart(chartSpace, None)

        # -- two plots under one plot area --
        assert len(chart.plots) == 2
        plot_kinds = {type(p).__name__ for p in chart.plots}
        assert plot_kinds == {"BarPlot", "LinePlot"}
        # -- secondary value axis is visible through the public API --
        assert chart.has_secondary_value_axis is True
        secondary = chart.secondary_value_axis
        assert isinstance(secondary, ValueAxis)
        # -- reads author-set scale on the secondary axis --
        assert secondary.maximum_scale == 100.0
        # -- the line plot's c:axId pair binds to the secondary valAx --
        lineChart = chart._chartSpace.plotArea.xCharts[1]
        line_axIds = [e.get("val") for e in lineChart.findall(qn("c:axId"))]
        assert line_axIds[1] == "333"
        assert secondary._element.find(qn("c:axId")).get("val") == "333"

    def it_allows_mutating_the_secondary_axis_on_a_combo_chart(self):
        """Secondary-axis scale/font writes persist to the right-hand valAx only (#141)."""
        chartSpace = parse_xml(_combo_chartSpace_with_secondary_axis_xml())
        chart = Chart(chartSpace, None)

        chart.secondary_value_axis.maximum_scale = 250.0

        primary, secondary = chartSpace.xpath(".//c:valAx")
        # -- write landed on the secondary axis --
        max_elms = secondary.xpath("./c:scaling/c:max")
        assert len(max_elms) == 1
        assert float(max_elms[0].get("val")) == 250.0
        # -- primary axis stayed untouched --
        assert primary.xpath("./c:scaling/c:max") == []

    # -- (3) custom formatting on the overlay plot --------------------

    def it_preserves_line_color_and_marker_style_across_save_reopen(self, _restore_part_factory):
        """Line color + marker style on an ``add_plot`` series survive round-trip."""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        bar_data = CategoryChartData()
        bar_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        bar_data.add_series("Revenue", (10, 20, 30, 40))
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            bar_data,
        ).chart

        line_data = CategoryChartData()
        line_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        line_data.add_series("Trend", (5, 15, 25, 35))
        line_plot = chart.add_plot(XL_CHART_TYPE.LINE_MARKERS, line_data)

        # -- apply per-series formatting that the reporter might want --
        trend_series = line_plot.series[0]
        trend_series.format.line.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        trend_series.marker.style = XL_MARKER_STYLE.CIRCLE
        trend_series.marker.size = 10

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        chart2 = next(s.chart for s in prs2.slides[0].shapes if s.has_chart)

        # -- locate the line series after reopen --
        line_plot2 = next(p for p in chart2.plots if isinstance(p, LinePlot))
        trend2 = line_plot2.series[0]
        assert trend2.name == "Trend"
        # -- color survived --
        assert trend2.format.line.color.rgb == RGBColor(0xFF, 0x00, 0x00)
        # -- marker style + size survived --
        assert trend2.marker.style == XL_MARKER_STYLE.CIRCLE
        assert trend2.marker.size == 10

    # -- reporter-workflow smoke test (the "is there any other way?" ask) --

    def it_does_not_raise_the_AttributeError_from_the_470_report(self, _restore_part_factory):
        """Pin that the ``add_plot`` path avoids the 'x_values_ref' crash the reporter hit.

        The #470 report is a crash from a pre-fork recipe that tried to
        reuse ``CategoryWorkbookWriter`` for a non-existent XY workbook
        writer. ``Chart.add_plot`` emits inline ``c:numLit`` / ``c:strLit``
        elements instead (see ``src/pptx/chart/xmlwriter.py``
        ``_PlotFragmentBuilder``), bypassing the embedded-workbook
        writer entirely. This test locks in the "no crash, chart
        saves cleanly" behaviour end-to-end.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        bar_data = CategoryChartData()
        bar_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        bar_data.add_series("Revenue", (10, 20, 30, 40))
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            bar_data,
        ).chart

        line_data = CategoryChartData()
        line_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        line_data.add_series("Trend", (5, 15, 25, 35))
        # -- pre-fork this path raised: --
        # --   AttributeError: 'CategoryWorkbookWriter' object has no --
        # --   attribute 'x_values_ref' --
        chart.add_plot(XL_CHART_TYPE.LINE, line_data)

        # -- the save path also must not raise --
        buf = io.BytesIO()
        prs.save(buf)
        assert buf.tell() > 0
        # -- inline numLit / strLit used on the overlay plot (no workbook --
        # -- rewrite for the second plot's data) --
        lineChart = chart._chartSpace.plotArea.xCharts[1]
        numLits = lineChart.findall(f".//{qn('c:numLit')}")
        strLits = lineChart.findall(f".//{qn('c:cat')}/{qn('c:strLit')}")
        assert len(numLits) >= 1
        assert len(strLits) == 1
