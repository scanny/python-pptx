# pyright: reportPrivateUsage=false

"""Regression test for issue #619 — add data labels to XY Scatter chart.

Issue #619 (https://github.com/scanny/python-pptx/issues/619) reports that
setting ``plot.has_data_labels = True`` on an XY scatter chart raises an
``AttributeError`` because the underlying ``CT_ScatterChart`` oxml class
was missing its ``dLbls`` ``ZeroOrOne`` descriptor. As a result the common
idiom::

    chart.plots[0].has_data_labels = True

worked on bar, line, pie, area, etc. but blew up on scatter charts.

This test exercises the end-to-end flow on all five XY scatter chart
types, including the save / reload round-trip, to lock in the fix and
guard against regression. The ``c:dLbls`` subtree emitted on scatter is
the same as on other plot types (``c:showLegendKey``, ``c:showVal``,
``c:showCatName``, ``c:showSerName``, ``c:showPercent``,
``c:showBubbleSize``, ``c:showLeaderLines``) — which is what we want,
since PowerPoint picks sensible per-plot-type defaults from the zeroed
flags.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.chart.data import XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches


def _build_scatter_chart(chart_type):
    """Build an XY scatter chart of *chart_type* and return the plot."""
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
    return prs, graphic_frame.chart.plots[0]


_ALL_SCATTER_TYPES = [
    XL_CHART_TYPE.XY_SCATTER,
    XL_CHART_TYPE.XY_SCATTER_LINES,
    XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
    XL_CHART_TYPE.XY_SCATTER_SMOOTH,
    XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
]


class DescribeIssue619XyScatterDataLabels(object):
    @pytest.mark.parametrize("chart_type", _ALL_SCATTER_TYPES)
    def it_starts_without_data_labels(self, chart_type):
        _, plot = _build_scatter_chart(chart_type)

        assert plot.has_data_labels is False

    @pytest.mark.parametrize("chart_type", _ALL_SCATTER_TYPES)
    def it_can_enable_data_labels(self, chart_type):
        _, plot = _build_scatter_chart(chart_type)

        plot.has_data_labels = True

        assert plot.has_data_labels is True
        # ---regression for #619: accessing .data_labels after toggling
        # has_data_labels must not raise AttributeError.
        assert plot.data_labels.show_value is True

    @pytest.mark.parametrize("chart_type", _ALL_SCATTER_TYPES)
    def it_can_disable_data_labels(self, chart_type):
        _, plot = _build_scatter_chart(chart_type)
        plot.has_data_labels = True
        assert plot.has_data_labels is True

        plot.has_data_labels = False

        assert plot.has_data_labels is False

    def it_round_trips_data_labels_through_save_reload(self):
        prs, plot = _build_scatter_chart(XL_CHART_TYPE.XY_SCATTER)
        plot.has_data_labels = True
        # tweak a flag that is distinguishable from the default so we know
        # the dLbls survived the round-trip and wasn't dropped by the writer.
        plot.data_labels.show_value = False
        plot.data_labels.show_series_name = True

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reloaded = Presentation(buf)

        reloaded_plot = None
        for shp in reloaded.slides[0].shapes:
            if shp.has_chart:
                reloaded_plot = shp.chart.plots[0]
                break
        assert reloaded_plot is not None
        assert reloaded_plot.has_data_labels is True
        assert reloaded_plot.data_labels.show_value is False
        assert reloaded_plot.data_labels.show_series_name is True

    def it_emits_dLbls_with_scatter_appropriate_flags(self):
        _, plot = _build_scatter_chart(XL_CHART_TYPE.XY_SCATTER)

        plot.has_data_labels = True

        dLbls = plot._element.dLbls
        assert dLbls is not None
        # ---scatter dLbls carry the same default-off show-flags as the other
        # Cartesian plots; PowerPoint picks a sensible default label position
        # when none is specified.
        assert dLbls.showLegendKey.val is False
        # `has_data_labels = True` flips showVal to on (the common case).
        assert dLbls.showVal.val is True
        assert dLbls.showCatName.val is False
        assert dLbls.showSerName.val is False
        assert dLbls.showPercent.val is False
        # c:showBubbleSize is only meaningful on bubble charts, but the default
        # XML keeps it present and zero for consistency with CT_DLbls.new_dLbls.
        assert dLbls.xpath("./c:showBubbleSize")[0].get("val") == "0"
