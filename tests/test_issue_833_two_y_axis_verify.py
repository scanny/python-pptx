# pyright: reportPrivateUsage=false

"""Regression test for issue #833 — 2 y-axis chart (secondary value axis).

Issue #833 (https://github.com/scanny/python-pptx/issues/833) asked for a
way to build a chart with two Y (value) axes — the familiar PowerPoint
"columns on the left axis, line on the right axis" pattern — so that two
data series on widely-different scales can be compared on the same chart.

The reporter's ask is satisfied on this fork by the combination of two
features that landed earlier in the release line:

* **#141 — secondary value axis read access.** Adds
  :attr:`Chart.has_secondary_value_axis` and
  :attr:`Chart.secondary_value_axis` so callers can detect and
  read/write the secondary ``c:valAx`` PowerPoint writes for
  category-based charts configured with a right-hand Y axis.

* **#470 — combo charts via Chart.add_plot.** Adds
  :meth:`Chart.add_plot` so a second plot (e.g. a line overlaid on a
  column chart) can be appended to an existing chart, sharing the
  category axis and optionally bound to a secondary value axis.

Together these deliver the #833 "two y-axis chart" use case: author a
combo chart, surface the secondary value axis, set independent
scale/major-unit on both axes, and round-trip through
:meth:`Presentation.save` + reopen.

Scope / known limitations
-------------------------
1. Authoring a secondary ``c:valAx`` element from scratch on a
   library-created chart (i.e. ``Chart.add_secondary_value_axis()`` or
   ``add_plot(..., secondary=True)``) is not directly exposed. The
   supported round-trip path is (a) open a PowerPoint-authored chart
   that already carries the secondary axis, or (b) author the
   ``chartSpace`` XML directly and mutate via
   :attr:`Chart.secondary_value_axis`.
2. :attr:`Chart.value_axis` is biased toward the XY/scatter case: when
   two ``c:valAx`` elements are present it returns the *second* one in
   document order. For a category-based combo chart that means
   ``chart.value_axis`` returns the same element as
   ``chart.secondary_value_axis``; the primary (left-hand) ``c:valAx``
   is only reachable via the XML plumbing
   (``chart._chartSpace.plotArea.primary_valAx``). This test suite
   pins the observed behaviour so a future refactor cannot change it
   silently.

See ``docs/user/charts.rst`` under "Secondary value axis" and
``docs/dev/analysis/combo-chart.rst``. #833 closes as delivered because
the reporter's core ask — "is there a way to build a chart with two y
axes?" — is satisfied end-to-end by #141 + #470.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.chart.axis import ValueAxis
from pptx.chart.chart import Chart
from pptx.chart.data import CategoryChartData
from pptx.chart.plot import BarPlot, LinePlot
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Inches

# -- test helpers --------------------------------------------------------


def _two_y_axis_chartSpace_xml():
    """Return chartSpace XML for a column+line combo with a secondary valAx.

    Shape mirrors what PowerPoint writes for the "columns on the primary
    axis, line on the secondary axis" combo chart (the #833 request):

    * ``c:barChart`` referencing ``(catAx=111, valAx=222)`` (primary left).
    * ``c:lineChart`` referencing ``(catAx=111, valAx=333)`` (secondary right).
    * three axes: shared ``c:catAx@axId=111``, primary
      ``c:valAx@axId=222`` on the left, secondary ``c:valAx@axId=333`` on
      the right.

    Neither value axis carries an author-set scale; the regression test
    exercises *writing* scale/major-unit on both from a clean baseline.
    """
    return (
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
        '/2006/chart">'
        "<c:chart><c:plotArea>"
        # -- back-most plot: column chart bound to primary value axis --
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
        # -- primary value axis (left) --
        "<c:valAx>"
        '<c:axId val="222"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/>'
        '<c:axPos val="l"/>'
        '<c:crossAx val="111"/>'
        "</c:valAx>"
        # -- secondary value axis (right) --
        "<c:valAx>"
        '<c:axId val="333"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/>'
        '<c:axPos val="r"/>'
        '<c:crossAx val="111"/>'
        "</c:valAx>"
        "</c:plotArea></c:chart></c:chartSpace>"
    )


# -- fixtures ------------------------------------------------------------


@pytest.fixture
def _restore_part_factory():
    """Guard against pollution from tests that mutate `PartFactory.part_type_for`.

    ``tests/opc/test_package.py::DescribePartFactory`` overwrites
    ``PartFactory.part_type_for[CT.PML_SLIDE]`` with a Mock and does
    not restore it. Mirrors the identical fixture elsewhere in the
    regression-test suite (``tests/test_issue_303_multi_axes_verify.py``,
    ``tests/test_issue_470_combo_charts_verify.py``).
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


# -- the verification suite ---------------------------------------------


class DescribeIssue833TwoYAxisVerify(object):
    """#833 two-y-axis verify-and-close via #141 + #470.

    Pins the reporter's workflow: author a combo chart (bar + line)
    with a secondary value axis, set independent scale/major-unit on
    both axes, and round-trip through save + reopen.
    """

    # -- authoring: combo chart with a secondary value axis ----------------

    def it_authors_a_bar_plus_line_combo_with_a_secondary_value_axis(self):
        """A two-y-axis combo is composable from pre-authored chartSpace XML (#141, #470)."""
        chartSpace = parse_xml(_two_y_axis_chartSpace_xml())
        chart = Chart(chartSpace, None)

        # -- two plots under one plot area --
        assert len(chart.plots) == 2
        plot_kinds = {type(p).__name__ for p in chart.plots}
        assert plot_kinds == {"BarPlot", "LinePlot"}
        assert isinstance(chart.plots[0], BarPlot)
        assert isinstance(chart.plots[1], LinePlot)
        # -- secondary value axis is visible through the public API --
        assert chart.has_secondary_value_axis is True
        assert isinstance(chart.secondary_value_axis, ValueAxis)
        # -- secondary axis binds to the second-in-document-order c:valAx --
        valAx_lst = chartSpace.xpath(".//c:valAx")
        assert chart.secondary_value_axis._element is valAx_lst[1]
        # -- the line plot's c:axId pair targets the secondary valAx --
        lineChart = chart._chartSpace.plotArea.xCharts[1]
        line_axIds = [e.get("val") for e in lineChart.findall(qn("c:axId"))]
        assert line_axIds[1] == "333"

    # -- independent scale + major_unit on both axes ----------------------

    def it_allows_setting_scale_and_major_unit_on_the_secondary_axis(self):
        """Writes on ``chart.secondary_value_axis`` land on the secondary ``c:valAx`` only (#141)."""
        chartSpace = parse_xml(_two_y_axis_chartSpace_xml())
        chart = Chart(chartSpace, None)

        chart.secondary_value_axis.minimum_scale = 0.0
        chart.secondary_value_axis.maximum_scale = 100.0
        chart.secondary_value_axis.major_unit = 25.0

        primary, secondary = chartSpace.xpath(".//c:valAx")
        # -- secondary scale/major-unit written --
        assert float(secondary.xpath("./c:scaling/c:min")[0].get("val")) == 0.0
        assert float(secondary.xpath("./c:scaling/c:max")[0].get("val")) == 100.0
        assert float(secondary.xpath("./c:majorUnit")[0].get("val")) == 25.0
        # -- primary axis stayed untouched --
        assert primary.xpath("./c:scaling/c:min") == []
        assert primary.xpath("./c:scaling/c:max") == []
        assert primary.xpath("./c:majorUnit") == []

    def it_allows_setting_scale_and_major_unit_on_the_primary_axis_via_xml(self):
        """The primary (left) axis is reachable via the plotArea element (#141)."""
        chartSpace = parse_xml(_two_y_axis_chartSpace_xml())
        chart = Chart(chartSpace, None)

        # -- primary valAx lives on the plotArea; wrap it as a ValueAxis --
        primary_valAx = chart._chartSpace.plotArea.primary_valAx
        assert primary_valAx is not None
        primary_axis = ValueAxis(primary_valAx)

        primary_axis.minimum_scale = 0.0
        primary_axis.maximum_scale = 50.0
        primary_axis.major_unit = 10.0

        primary, secondary = chartSpace.xpath(".//c:valAx")
        # -- primary scale/major-unit written --
        assert float(primary.xpath("./c:scaling/c:min")[0].get("val")) == 0.0
        assert float(primary.xpath("./c:scaling/c:max")[0].get("val")) == 50.0
        assert float(primary.xpath("./c:majorUnit")[0].get("val")) == 10.0
        # -- secondary axis stayed untouched --
        assert secondary.xpath("./c:scaling/c:min") == []
        assert secondary.xpath("./c:scaling/c:max") == []
        assert secondary.xpath("./c:majorUnit") == []

    def it_keeps_primary_and_secondary_axis_writes_independent(self):
        """Writing min/max/major_unit on one axis never bleeds onto the other (#141)."""
        chartSpace = parse_xml(_two_y_axis_chartSpace_xml())
        chart = Chart(chartSpace, None)

        primary_axis = ValueAxis(chart._chartSpace.plotArea.primary_valAx)

        # -- write different scales to each axis --
        primary_axis.minimum_scale = 0.0
        primary_axis.maximum_scale = 50.0
        primary_axis.major_unit = 10.0
        chart.secondary_value_axis.minimum_scale = 0.0
        chart.secondary_value_axis.maximum_scale = 100.0
        chart.secondary_value_axis.major_unit = 25.0

        # -- reads round-trip through each axis's own proxy --
        assert primary_axis.minimum_scale == 0.0
        assert primary_axis.maximum_scale == 50.0
        assert primary_axis.major_unit == 10.0
        assert chart.secondary_value_axis.minimum_scale == 0.0
        assert chart.secondary_value_axis.maximum_scale == 100.0
        assert chart.secondary_value_axis.major_unit == 25.0

    # -- chart.value_axis quirk on combo charts --------------------------

    def it_returns_the_secondary_valAx_from_chart_value_axis_on_a_combo(self):
        """`chart.value_axis` aliases the second `c:valAx` when two are present.

        Documented quirk inherited from the XY/scatter case (where the
        second ``c:valAx`` is the Y-axis). For a category-based combo
        chart, callers wanting the primary (left-hand) axis should use
        :attr:`Chart.primary_value_axis` (FU-2); the XML-level escape
        hatch via ``chart._chartSpace.plotArea.primary_valAx`` also
        remains available.
        """
        chartSpace = parse_xml(_two_y_axis_chartSpace_xml())
        chart = Chart(chartSpace, None)

        valAx_lst = chartSpace.xpath(".//c:valAx")
        assert chart.value_axis._element is valAx_lst[1]
        assert chart.value_axis._element is chart.secondary_value_axis._element

    def it_returns_the_primary_valAx_from_primary_value_axis_on_a_combo(self):
        """`chart.primary_value_axis` resolves to the FIRST `c:valAx` (FU-2).

        The backward-compatible fix for the `Chart.value_axis` combo-chart
        quirk pinned above: add `Chart.primary_value_axis` returning the
        first ``c:valAx`` unambiguously. Demonstrated here on the same
        combo-chart fixture used throughout this module.
        """
        chartSpace = parse_xml(_two_y_axis_chartSpace_xml())
        chart = Chart(chartSpace, None)

        valAx_lst = chartSpace.xpath(".//c:valAx")
        primary = chart.primary_value_axis

        assert isinstance(primary, ValueAxis)
        assert primary._element is valAx_lst[0]
        # -- and explicitly not aliased to secondary --
        assert primary._element is not chart.secondary_value_axis._element
        assert primary._element is not chart.value_axis._element

    # -- round-trip save/reopen ------------------------------------------

    def it_round_trips_independent_axis_scales_through_save_and_reopen(
        self, _restore_part_factory
    ):
        """Independent primary/secondary scale+major_unit survive a save+reopen cycle."""
        # -- start from a fresh bar chart on a real slide --
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        bar_data = CategoryChartData()
        bar_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        bar_data.add_series("Revenue", (10, 20, 30, 40))
        graphic_frame = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            bar_data,
        )
        chart = graphic_frame.chart

        # -- overlay a line plot so both series share the category axis --
        line_data = CategoryChartData()
        line_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        line_data.add_series("Trend", (5, 15, 25, 35))
        chart.add_plot(XL_CHART_TYPE.LINE, line_data)

        # -- splice a second `c:valAx` into the plotArea and re-wire the --
        # -- line chart's second c:axId to target the new axis. This is --
        # -- the documented "author the XML directly" path called out in --
        # -- docs/user/charts.rst — Chart.add_plot does not yet expose --
        # -- `secondary=True`, so we rewrite the element here. --
        plotArea = chart._chartSpace.plotArea
        primary_valAx = plotArea.valAx_lst[0]
        primary_axId_val = primary_valAx.find(qn("c:axId")).get("val")
        # -- pick a fresh id distinct from the existing ones --
        existing_ids = {
            e.get("val")
            for e in chart._chartSpace.iter(qn("c:axId"))
        }
        secondary_axId_val = "98765"
        assert secondary_axId_val not in existing_ids
        secondary_valAx_xml = (
            '<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            f'<c:axId val="{secondary_axId_val}"/>'
            '<c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/>'
            '<c:axPos val="r"/>'
            f'<c:crossAx val="{primary_axId_val}"/>'
            "</c:valAx>"
        )
        secondary_valAx = parse_xml(secondary_valAx_xml)
        primary_valAx.addnext(secondary_valAx)
        # -- rewire the line chart's value-axis pointer onto the new axis --
        lineChart = plotArea.xCharts[1]
        line_axIds = lineChart.findall(qn("c:axId"))
        line_axIds[1].set("val", secondary_axId_val)

        # -- set independent scale + major_unit on both axes --
        primary_axis = ValueAxis(plotArea.primary_valAx)
        primary_axis.minimum_scale = 0.0
        primary_axis.maximum_scale = 50.0
        primary_axis.major_unit = 10.0
        chart.secondary_value_axis.minimum_scale = 0.0
        chart.secondary_value_axis.maximum_scale = 100.0
        chart.secondary_value_axis.major_unit = 25.0

        # -- round-trip --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        chart2 = next(s.chart for s in prs2.slides[0].shapes if s.has_chart)

        # -- both plot types present on reopen --
        assert len(chart2.plots) == 2
        plot_kinds2 = {type(p).__name__ for p in chart2.plots}
        assert plot_kinds2 == {"BarPlot", "LinePlot"}
        # -- secondary axis survives round-trip --
        assert chart2.has_secondary_value_axis is True
        # -- primary scale + major_unit round-tripped --
        primary_axis2 = ValueAxis(chart2._chartSpace.plotArea.primary_valAx)
        assert primary_axis2.minimum_scale == 0.0
        assert primary_axis2.maximum_scale == 50.0
        assert primary_axis2.major_unit == 10.0
        # -- secondary scale + major_unit round-tripped independently --
        assert chart2.secondary_value_axis.minimum_scale == 0.0
        assert chart2.secondary_value_axis.maximum_scale == 100.0
        assert chart2.secondary_value_axis.major_unit == 25.0

    # -- negative cases ---------------------------------------------------

    def it_reports_False_for_has_secondary_on_a_plain_single_axis_chart(self):
        """A single-valAx chart correctly reports no secondary value axis (#833)."""
        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea>"
            "<c:barChart>"
            '<c:barDir val="col"/><c:grouping val="clustered"/>'
            '<c:axId val="1"/><c:axId val="2"/>'
            "</c:barChart>"
            '<c:catAx><c:axId val="1"/></c:catAx>'
            '<c:valAx><c:axId val="2"/></c:valAx>'
            "</c:plotArea></c:chart></c:chartSpace>"
        )
        chart = Chart(parse_xml(chartSpace_xml), None)

        assert chart.has_secondary_value_axis is False
        with pytest.raises(ValueError, match="no secondary value axis"):
            chart.secondary_value_axis
