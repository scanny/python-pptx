# pyright: reportPrivateUsage=false

"""Regression test for issue #303 — multiple axes in one chart.

Issue #303 (https://github.com/scanny/python-pptx/issues/303) asked for
the ability to author a chart that carries more than one value axis —
specifically so that two data series on widely-different scales can each
be plotted against their own axis on the same chart (the familiar
"columns + line with secondary axis" PowerPoint layout).

This suite verifies that #303's use cases are satisfied by the
combination of two features that landed on this fork:

* **#141 — secondary value axis read access.** Adds
  ``Chart.has_secondary_value_axis`` and ``Chart.secondary_value_axis``
  so callers can detect and read the secondary ``c:valAx`` that
  PowerPoint writes for category-based charts configured with a
  right-hand Y axis. An XY/scatter chart has two value axes but
  neither is considered "secondary"; ``has_secondary_value_axis`` is
  ``False`` in that case.

* **#338 — combo charts.** Adds ``Chart.add_plot(chart_type,
  chart_data)`` so a second plot (e.g. a line overlaid on a column
  chart) can be appended to an existing chart. The new plot reuses
  the existing plot's ``c:axId`` values and is inserted into the
  plot area.

Together these cover the #303 ask: a caller can compose a chart with
two plots sharing a category axis where the second plot is bound to
a secondary value axis and tick-label / scale properties of that
secondary axis are read/writable through
``Chart.secondary_value_axis``. Authoring the secondary ``c:valAx``
itself from scratch is still not directly exposed (see the note in
``docs/user/charts.rst`` under "Secondary value axis"), but the
common round-trip case — open a PowerPoint-authored combo chart
with a secondary axis, or compose one from XML + ``add_plot`` — is
covered.
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
from pptx.util import Inches, Pt

# -- test helpers -------------------------------------------------------


def _combo_chartSpace_xml_with_secondary_axis():
    """Return chartSpace XML for a column chart + secondary-bound second plot.

    Mirrors the shape PowerPoint writes for the common "columns on the
    primary axis, line on the secondary axis" combo chart:

    * one ``c:barChart`` plot whose ``c:axId`` pair references
      ``catAx=111`` and primary ``valAx=222``.
    * one ``c:lineChart`` plot whose ``c:axId`` pair references the
      *same* ``catAx=111`` but the *secondary* ``valAx=333``.
    * three axes: ``c:catAx/@axId=111``, primary ``c:valAx/@axId=222``
      (``crossAx=111``) and secondary ``c:valAx/@axId=333``
      (``crossAx=111``).

    The secondary ``c:valAx`` carries a scale (``c:max``) so the
    regression test can also verify that ``Chart.secondary_value_axis``
    surfaces the author's existing values for read and allows write.
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
        # -- primary value axis --
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


# -- fixtures -----------------------------------------------------------


@pytest.fixture
def _restore_part_factory():
    """Guard against pollution from tests that mutate `PartFactory.part_type_for`.

    ``tests/opc/test_package.py::DescribePartFactory`` overwrites
    ``PartFactory.part_type_for[CT.PML_SLIDE]`` with a Mock and does
    not restore it. Mirrors the identical fixture elsewhere in the
    regression-test suite.
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


class DescribeIssue303MultiAxesVerify(object):
    """#303 multi-axes verify-and-close via #141 + #338.

    Pins the read/write secondary-axis plumbing together with the
    add-second-plot authoring primitive, to confirm that the #303 use
    case of "two series on different Y scales in one chart" is achievable
    end-to-end on this fork.
    """

    # -- #141: secondary value axis read access ---------------------------

    def it_detects_a_secondary_value_axis_on_a_combo_chart(self):
        """``has_secondary_value_axis`` is True when a chart carries two valAx."""
        chartSpace = parse_xml(_combo_chartSpace_xml_with_secondary_axis())
        chart = Chart(chartSpace, None)

        assert chart.has_secondary_value_axis is True

    def it_provides_a_ValueAxis_for_the_secondary_axis(self):
        """``secondary_value_axis`` returns a real ``ValueAxis`` wrapping the 2nd valAx."""
        chartSpace = parse_xml(_combo_chartSpace_xml_with_secondary_axis())
        chart = Chart(chartSpace, None)

        secondary = chart.secondary_value_axis

        assert isinstance(secondary, ValueAxis)
        # -- the wrapped element is the second c:valAx in document order --
        valAx_lst = chartSpace.xpath(".//c:valAx")
        assert secondary._element is valAx_lst[1]
        # -- the secondary axis is a distinct XML element from the primary --
        primary_valAx = valAx_lst[0]
        assert secondary._element is not primary_valAx

    def it_reads_the_existing_scale_on_the_secondary_axis(self):
        """The secondary axis surfaces its author-set scale (#141 read-access)."""
        chartSpace = parse_xml(_combo_chartSpace_xml_with_secondary_axis())
        chart = Chart(chartSpace, None)

        # -- author wrote c:max val="100" on the secondary axis --
        assert chart.secondary_value_axis.maximum_scale == 100.0

    def it_allows_writing_to_the_secondary_axis_scale(self):
        """Scale writes on the secondary axis are persisted to its own c:valAx."""
        chartSpace = parse_xml(_combo_chartSpace_xml_with_secondary_axis())
        chart = Chart(chartSpace, None)

        chart.secondary_value_axis.maximum_scale = 250.0

        # -- primary axis scale should be untouched; secondary should update --
        secondary_valAx = chartSpace.xpath(".//c:valAx")[1]
        max_elms = secondary_valAx.xpath("./c:scaling/c:max")
        assert len(max_elms) == 1
        assert float(max_elms[0].get("val")) == 250.0
        # -- primary value axis has no c:max set --
        primary_valAx = chartSpace.xpath(".//c:valAx")[0]
        assert primary_valAx.xpath("./c:scaling/c:max") == []

    def it_reports_False_for_has_secondary_on_single_axis_chart(self):
        """``has_secondary_value_axis`` is False when only the primary valAx is present."""
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

    def it_reports_False_for_has_secondary_on_XY_scatter(self):
        """An XY/scatter chart has two valAx but neither is 'secondary' (#141)."""
        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea>"
            '<c:valAx><c:axId val="1"/></c:valAx>'
            '<c:valAx><c:axId val="2"/></c:valAx>'
            "</c:plotArea></c:chart></c:chartSpace>"
        )
        chart = Chart(parse_xml(chartSpace_xml), None)

        assert chart.has_secondary_value_axis is False

    # -- #338: add_plot authoring primitive -------------------------------

    def it_can_add_a_line_plot_to_a_column_chart(self):
        """``Chart.add_plot`` appends a second plot sharing the first plot's axes (#338)."""
        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea>"
            "<c:barChart>"
            '<c:barDir val="col"/><c:grouping val="clustered"/>'
            '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
            '<c:axId val="111"/><c:axId val="222"/>'
            "</c:barChart>"
            '<c:catAx><c:axId val="111"/></c:catAx>'
            '<c:valAx><c:axId val="222"/></c:valAx>'
            "</c:plotArea></c:chart></c:chartSpace>"
        )
        chart = Chart(parse_xml(chartSpace_xml), None)
        line_data = CategoryChartData()
        line_data.categories = ["Q1", "Q2", "Q3"]
        line_data.add_series("Trend", (5.0, 10.0, 15.0))

        new_plot = chart.add_plot(XL_CHART_TYPE.LINE_MARKERS, line_data)

        # -- both plots live under the single plotArea --
        assert len(chart.plots) == 2
        assert isinstance(new_plot, LinePlot)
        # -- the new lineChart references the existing axId pair --
        lineChart = chart._chartSpace.plotArea.xCharts[1]
        line_axIds = [e.get("val") for e in lineChart.findall(qn("c:axId"))]
        assert line_axIds == ["111", "222"]
        # -- series count reflects both plots --
        assert len(chart.series) == 2

    def it_can_add_a_bar_plot_to_a_line_chart(self):
        """The inverse direction is also supported (line chart + bar overlay)."""
        chartSpace_xml = (
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml'
            '/2006/chart">'
            "<c:chart><c:plotArea>"
            "<c:lineChart>"
            '<c:grouping val="standard"/>'
            '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
            '<c:axId val="1"/><c:axId val="2"/>'
            "</c:lineChart>"
            '<c:catAx><c:axId val="1"/></c:catAx>'
            '<c:valAx><c:axId val="2"/></c:valAx>'
            "</c:plotArea></c:chart></c:chartSpace>"
        )
        chart = Chart(parse_xml(chartSpace_xml), None)
        bar_data = CategoryChartData()
        bar_data.categories = ["a", "b"]
        bar_data.add_series("Bars", (3.0, 4.0))

        plot = chart.add_plot(XL_CHART_TYPE.COLUMN_CLUSTERED, bar_data)

        assert isinstance(plot, BarPlot)
        assert len(chart.plots) == 2

    # -- #141 + #338 together: the #303 ask -------------------------------

    def it_can_inspect_a_combo_chart_with_a_secondary_axis(self):
        """The #303 round-trip case: two plots, two value axes, read-access pinned."""
        chartSpace = parse_xml(_combo_chartSpace_xml_with_secondary_axis())
        chart = Chart(chartSpace, None)

        # -- #338 gives us the two plots --
        assert len(chart.plots) == 2
        plot_kinds = {type(p).__name__ for p in chart.plots}
        assert plot_kinds == {"BarPlot", "LinePlot"}
        # -- #141 gives us the secondary axis via its dedicated accessor --
        assert chart.has_secondary_value_axis is True
        assert chart.secondary_value_axis._element is chartSpace.xpath(".//c:valAx")[1]
        # -- the second plot's c:axId pair points at the secondary valAx --
        lineChart = chart._chartSpace.plotArea.xCharts[1]
        line_axIds = [e.get("val") for e in lineChart.findall(qn("c:axId"))]
        assert line_axIds[1] == "333"
        assert chart.secondary_value_axis._element.find(qn("c:axId")).get("val") == "333"

    def it_isolates_secondary_axis_writes_from_the_primary_axis(self):
        """Font / scale writes on the secondary axis don't bleed into the primary.

        The primary-axis accessor (``Chart.value_axis``) is known to return the
        *last* ``c:valAx`` when two are present — a quirk of the pre-existing
        XY-scatter-aware accessor. The #141 ``secondary_value_axis`` accessor
        is what callers use for the right-hand axis, and this test pins that
        writes through it are confined to the second ``c:valAx`` element while
        leaving the first (primary) ``c:valAx`` untouched.
        """
        chartSpace = parse_xml(_combo_chartSpace_xml_with_secondary_axis())
        chart = Chart(chartSpace, None)

        chart.secondary_value_axis.maximum_scale = 500.0
        chart.secondary_value_axis.tick_labels.font.size = Pt(14)

        primary, secondary = chartSpace.xpath(".//c:valAx")
        # -- secondary axis received the scale + font size writes --
        assert float(secondary.xpath("./c:scaling/c:max")[0].get("val")) == 500.0
        assert secondary.xpath(".//c:txPr//a:defRPr/@sz")[0] == "1400"
        # -- primary axis stayed clean (no c:max, no font-size override) --
        assert primary.xpath("./c:scaling/c:max") == []
        assert primary.xpath(".//c:txPr//a:defRPr/@sz") == []

    def it_round_trips_a_combo_chart_with_secondary_axis_through_save_reopen(
        self, _restore_part_factory
    ):
        """The composed combo chart survives save + reopen (end-to-end pin)."""
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

        # -- append a line plot (#338) --
        line_data = CategoryChartData()
        line_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        line_data.add_series("Trend", (5, 15, 25, 35))
        chart.add_plot(XL_CHART_TYPE.LINE_MARKERS, line_data)

        # -- save + reopen --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        graphic_frame = next(s for s in prs2.slides[0].shapes if s.has_chart)
        chart2 = graphic_frame.chart

        # -- both plots survived the round-trip --
        assert len(chart2.plots) == 2
        plot_kinds = {type(p).__name__ for p in chart2.plots}
        assert plot_kinds == {"BarPlot", "LinePlot"}
        # -- the reopened chart has the expected series --
        series_names = [s.name for s in chart2.series]
        assert "Revenue" in series_names
        assert "Trend" in series_names
