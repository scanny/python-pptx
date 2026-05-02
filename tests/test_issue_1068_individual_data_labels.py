# pyright: reportPrivateUsage=false

"""Regression test for issue #1068 — edit individual data labels.

Issue #1068 (https://github.com/scanny/python-pptx/issues/1068) asks
whether it is possible to edit *individual* data labels rather than
the whole ``series.data_labels`` collection — toggling ``show_value`` /
``show_series_name``, customizing text, position, font, and fill for a
single data point.

This capability has always been available through
``series.points[i].data_label`` (the |DataLabel| per-point proxy
wrapping the ``c:dLbl`` element addressed by a ``c:idx`` matching the
point index). The Wave-era work that closed the adjacent gaps —
``feat/issue-716-datalabel-border`` (Wave 1) added
``DataLabel.format``, and ``feat/issue-560-data-label-colors`` (Wave 7)
added the series-level analogue ``DataLabels.format`` — round out the
per-label API. What the #1068 reporter wanted (toggle label contents on
one point only) is achieved by assigning a ``text_frame.text`` on that
point's |DataLabel|, which writes a ``c:dLbl/c:tx/c:rich`` subtree that
PowerPoint renders *in place of* the series-level show-* toggles for
that point only — this is the PowerPoint idiom for "different label on
one point".

This module exercises the reporter's exact intent end-to-end and pins
the per-point independence property (customizing label N does not
touch label M) across |Series| families that expose ``points``
(CategoryPoints for |BarPlot| / |ColumnPlot| / |LinePlot|, XyPoints for
|XySeries|, BubblePoints for |BubbleSeries|).
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.chart.data import BubbleChartData, CategoryChartData, XyChartData
from pptx.chart.datalabel import DataLabel
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.util import Inches


def _category_chart():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    cd = CategoryChartData()
    cd.categories = ["A", "B", "C"]
    cd.add_series("s1", (1.0, 2.0, 3.0))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        cd,
    ).chart
    chart.plots[0].has_data_labels = True
    return prs, chart


def _xy_chart():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    cd = XyChartData()
    s = cd.add_series("s1")
    s.add_data_point(1.0, 2.0)
    s.add_data_point(2.0, 3.0)
    s.add_data_point(3.0, 4.0)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.XY_SCATTER,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        cd,
    ).chart
    chart.plots[0].has_data_labels = True
    return prs, chart


def _bubble_chart():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    cd = BubbleChartData()
    s = cd.add_series("s1")
    s.add_data_point(1.0, 2.0, 10.0)
    s.add_data_point(2.0, 3.0, 20.0)
    s.add_data_point(3.0, 4.0, 30.0)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BUBBLE,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        cd,
    ).chart
    chart.plots[0].has_data_labels = True
    return prs, chart


def _roundtrip(prs):
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


def _reloaded_series(prs):
    for shp in prs.slides[0].shapes:
        if shp.has_chart:
            return shp.chart.plots[0].series[0]
    raise AssertionError("reloaded presentation contains no chart")


class DescribeIssue1068IndividualDataLabels(object):
    """Verify #1068 — edit individual data labels — is resolved."""

    def it_exposes_a_DataLabel_per_point_on_category_series(self):
        """Reporter's core question: can I address a single label?

        ``series.points[i].data_label`` must return a |DataLabel| for
        every point in the series.
        """
        _, chart = _category_chart()
        series = chart.plots[0].series[0]

        labels = [p.data_label for p in series.points]

        assert len(labels) == 3
        assert all(isinstance(dl, DataLabel) for dl in labels)
        # ---distinct proxies, each pointing at its own c:idx---
        assert labels[0]._idx == 0
        assert labels[1]._idx == 1
        assert labels[2]._idx == 2

    def it_exposes_a_DataLabel_per_point_on_xy_series(self):
        _, chart = _xy_chart()
        series = chart.plots[0].series[0]

        labels = [p.data_label for p in series.points]

        assert len(labels) == 3
        assert all(isinstance(dl, DataLabel) for dl in labels)

    def it_exposes_a_DataLabel_per_point_on_bubble_series(self):
        _, chart = _bubble_chart()
        series = chart.plots[0].series[0]

        labels = [p.data_label for p in series.points]

        assert len(labels) == 3
        assert all(isinstance(dl, DataLabel) for dl in labels)

    def it_round_trips_custom_text_on_a_single_data_label(self):
        """The PowerPoint idiom for #1068's screenshot.

        Setting ``data_label.text_frame.text`` on one point writes
        ``c:dLbl/c:tx/c:rich`` for that point only — PowerPoint then
        renders that literal text instead of honoring the series-level
        ``show_value`` / ``show_series_name`` flags for that point.
        """
        prs, chart = _category_chart()
        series = chart.plots[0].series[0]

        series.points[1].data_label.text_frame.text = "highlight"

        reloaded = _reloaded_series(_roundtrip(prs))
        assert reloaded.points[1].data_label.has_text_frame is True
        assert reloaded.points[1].data_label.text_frame.text == "highlight"
        # ---other points left untouched---
        assert reloaded.points[0].data_label.has_text_frame is False
        assert reloaded.points[2].data_label.has_text_frame is False

    def it_round_trips_per_point_position_override(self):
        """Individual labels can override the series-level position."""
        prs, chart = _category_chart()
        series = chart.plots[0].series[0]

        series.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        series.points[0].data_label.position = XL_LABEL_POSITION.INSIDE_END

        reloaded = _reloaded_series(_roundtrip(prs))
        assert reloaded.data_labels.position == XL_LABEL_POSITION.OUTSIDE_END
        assert reloaded.points[0].data_label.position == XL_LABEL_POSITION.INSIDE_END
        # ---point without an explicit position inherits series default (None at point scope)---
        assert reloaded.points[1].data_label.position is None

    def it_round_trips_per_point_font_color(self):
        """Already covered by #560 — cross-verified here for #1068."""
        prs, chart = _category_chart()
        series = chart.plots[0].series[0]

        series.points[0].data_label.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        series.points[2].data_label.font.color.rgb = RGBColor(0x00, 0xFF, 0x00)

        reloaded = _reloaded_series(_roundtrip(prs))
        assert reloaded.points[0].data_label.font.color.rgb == RGBColor(0xFF, 0x00, 0x00)
        assert reloaded.points[2].data_label.font.color.rgb == RGBColor(0x00, 0xFF, 0x00)

    def it_round_trips_per_point_fill_independently_of_other_points(self):
        """#716 shipped DataLabel.format — pin per-point independence.

        Setting fill on point 1 must not bleed onto points 0 or 2.
        """
        prs, chart = _category_chart()
        series = chart.plots[0].series[0]

        series.points[1].data_label.format.fill.solid()
        series.points[1].data_label.format.fill.fore_color.rgb = RGBColor(0xAB, 0xCD, 0xEF)

        reloaded = _reloaded_series(_roundtrip(prs))
        assert reloaded.points[1].data_label.format.fill.fore_color.rgb == RGBColor(
            0xAB, 0xCD, 0xEF
        )
        # ---sibling points have no c:spPr customization---
        spPr_qn = "{http://schemas.openxmlformats.org/drawingml/2006/chart}spPr"
        p0_dLbl = reloaded.points[0].data_label._dLbl
        p2_dLbl = reloaded.points[2].data_label._dLbl
        if p0_dLbl is not None:
            assert p0_dLbl.find(spPr_qn) is None
        if p2_dLbl is not None:
            assert p2_dLbl.find(spPr_qn) is None

    def it_writes_one_c_dLbl_per_customized_point_with_matching_c_idx(self):
        """Pin the XML shape the Wave-era APIs were designed to produce.

        Two customized points must result in two ``c:dLbl`` children
        under ``c:dLbls``, each with a matching ``c:idx/@val``.
        """
        prs, chart = _category_chart()
        series = chart.plots[0].series[0]

        series.points[0].data_label.text_frame.text = "A"
        series.points[2].data_label.text_frame.text = "C"

        dLbls = series.data_labels._element
        chart_ns = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        dLbl_list = dLbls.findall("%sdLbl" % chart_ns)
        assert len(dLbl_list) == 2
        idx_values = sorted(int(d.find("%sidx" % chart_ns).get("val")) for d in dLbl_list)
        assert idx_values == [0, 2]

        # ---also round-trips---
        reloaded = _reloaded_series(_roundtrip(prs))
        assert reloaded.points[0].data_label.text_frame.text == "A"
        assert reloaded.points[1].data_label.has_text_frame is False
        assert reloaded.points[2].data_label.text_frame.text == "C"

    def it_round_trips_custom_text_on_an_xy_point(self):
        """The per-point data_label path works on XY scatter series too."""
        prs, chart = _xy_chart()
        series = chart.plots[0].series[0]

        series.points[2].data_label.text_frame.text = "peak"

        reloaded = _reloaded_series(_roundtrip(prs))
        assert reloaded.points[2].data_label.text_frame.text == "peak"

    def it_round_trips_custom_text_on_a_bubble_point(self):
        """And on bubble series."""
        prs, chart = _bubble_chart()
        series = chart.plots[0].series[0]

        series.points[0].data_label.text_frame.text = "first"

        reloaded = _reloaded_series(_roundtrip(prs))
        assert reloaded.points[0].data_label.text_frame.text == "first"
