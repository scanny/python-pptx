# pyright: reportPrivateUsage=false

"""Regression test for issue #825 — change points color in scatter plot.

Issue #825 (https://github.com/scanny/python-pptx/issues/825) reports that
users can't set the marker color of individual data points in an XY scatter
plot. The confusion is that ``point.format.fill`` writes ``c:dPt/c:spPr``,
which PowerPoint renders as the data-point's *line* fill rather than as the
marker fill — so on scatter charts (where points are drawn as markers) the
authored color appears to do nothing.

The correct path on scatter/line/radar (line-type) charts is
``point.marker.format.fill``, which writes to ``c:dPt/c:marker/c:spPr`` —
exactly the element PowerPoint consults for the marker swatch. That API has
existed since ``Point.marker`` was first added (mirroring the series-level
``.marker`` on |XySeries| / |LineSeries| / |RadarSeries|), and follows the
same pattern as the #357 pie-slice color fix (``series.points[i].format``).

This test locks in the end-to-end flow on all five XY scatter chart types,
including the save-reload round-trip, so any future regression of the
per-point marker-fill path would be caught here.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.chart.data import XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against pollution from tests that mutate `PartFactory.part_type_for`.

    Matches the pattern used in ``tests/test_issue_619_xy_scatter_data_labels.py``.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.comments import CommentAuthorsPart, CommentsPart
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
        CT.PML_COMMENTS: CommentsPart,
        CT.PML_COMMENT_AUTHORS: CommentAuthorsPart,
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


def _build_scatter_chart(chart_type):
    """Build an XY scatter chart of *chart_type* and return ``(prs, chart)``."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = XyChartData()
    series_1 = chart_data.add_series("Series 1")
    series_1.add_data_point(0.7, 2.7)
    series_1.add_data_point(1.8, 3.2)
    series_1.add_data_point(2.6, 0.8)
    graphic_frame = slide.shapes.add_chart(
        chart_type,
        Inches(2),
        Inches(2),
        Inches(6),
        Inches(4),
        chart_data,
    )
    return prs, graphic_frame.chart


_SCATTER_TYPES_WITH_MARKERS = [
    XL_CHART_TYPE.XY_SCATTER,
    XL_CHART_TYPE.XY_SCATTER_LINES,
    XL_CHART_TYPE.XY_SCATTER_SMOOTH,
]


class DescribeIssue825ScatterPointColors(object):
    @pytest.mark.parametrize("chart_type", _SCATTER_TYPES_WITH_MARKERS)
    def it_writes_per_point_marker_fill_on_dPt_marker_spPr(self, chart_type):
        _, chart = _build_scatter_chart(chart_type)
        series = chart.plots[0].series[0]

        # ---color the second data point red via point.marker.format.fill ---
        point = series.points[1]
        point.marker.format.fill.solid()
        point.marker.format.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)

        # ---the per-point color lives on c:dPt/c:marker/c:spPr (NOT c:dPt/c:spPr). ---
        ser = series._element
        dPt_list = ser.findall(qn("c:dPt"))
        assert len(dPt_list) == 1
        dPt = dPt_list[0]
        assert dPt.find(qn("c:idx")).get("val") == "1"
        srgb_vals = dPt.xpath("c:marker/c:spPr/a:solidFill/a:srgbClr/@val")
        assert srgb_vals == ["FF0000"]
        # ---c:dPt/c:spPr must NOT be emitted: that path is for non-marker series
        # ---(bar/column fill) and PowerPoint renders it as the data-point line,
        # ---not the marker color, on scatter charts.
        assert dPt.find(qn("c:spPr")) is None

    def it_supports_independent_colors_on_multiple_points(self):
        _, chart = _build_scatter_chart(XL_CHART_TYPE.XY_SCATTER)
        series = chart.plots[0].series[0]

        red = RGBColor(0xFF, 0x00, 0x00)
        green = RGBColor(0x00, 0xFF, 0x00)
        blue = RGBColor(0x00, 0x00, 0xFF)
        for point, rgb in zip(series.points, (red, green, blue)):
            point.marker.format.fill.solid()
            point.marker.format.fill.fore_color.rgb = rgb

        ser = series._element
        dPt_elements = ser.findall(qn("c:dPt"))
        assert len(dPt_elements) == 3
        # ---each c:dPt addresses a different point by c:idx and carries its
        # own marker fill; no cross-contamination. ---
        idx_to_rgb = {
            dPt.find(qn("c:idx")).get("val"): dPt.xpath(
                "c:marker/c:spPr/a:solidFill/a:srgbClr/@val"
            )[0]
            for dPt in dPt_elements
        }
        assert idx_to_rgb == {"0": "FF0000", "1": "00FF00", "2": "0000FF"}

    def it_round_trips_per_point_marker_fill_through_save_reload(
        self, _restore_part_factory
    ):
        prs, chart = _build_scatter_chart(XL_CHART_TYPE.XY_SCATTER)
        series = chart.plots[0].series[0]
        series.points[0].marker.format.fill.solid()
        series.points[0].marker.format.fill.fore_color.rgb = RGBColor(0xAB, 0xCD, 0xEF)
        series.points[2].marker.format.fill.solid()
        series.points[2].marker.format.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reloaded = Presentation(buf)

        reloaded_chart = None
        for shp in reloaded.slides[0].shapes:
            if shp.has_chart:
                reloaded_chart = shp.chart
                break
        assert reloaded_chart is not None
        reloaded_series = reloaded_chart.plots[0].series[0]
        assert (
            reloaded_series.points[0].marker.format.fill.fore_color.rgb
            == RGBColor(0xAB, 0xCD, 0xEF)
        )
        assert (
            reloaded_series.points[2].marker.format.fill.fore_color.rgb
            == RGBColor(0x12, 0x34, 0x56)
        )

    def it_reuses_the_same_dPt_when_recoloring_a_point(self):
        """Regression: re-assigning marker.fill on the same point must not
        create a second ``c:dPt`` with the same ``c:idx`` — the existing one
        is located by xpath and reused (see ``get_or_add_dPt_for_point``)."""
        _, chart = _build_scatter_chart(XL_CHART_TYPE.XY_SCATTER)
        series = chart.plots[0].series[0]

        point = series.points[0]
        point.marker.format.fill.solid()
        point.marker.format.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)
        point.marker.format.fill.fore_color.rgb = RGBColor(0x00, 0xFF, 0x00)

        ser = series._element
        dPt_list = ser.findall(qn("c:dPt"))
        assert len(dPt_list) == 1
        srgb_vals = dPt_list[0].xpath("c:marker/c:spPr/a:solidFill/a:srgbClr/@val")
        assert srgb_vals == ["00FF00"]
