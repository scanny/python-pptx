# pyright: reportPrivateUsage=false

"""Regression test for issue #662 — data-label background color / fill.

Issue #662 (https://github.com/scanny/python-pptx/issues/662) asks for
a way to set the *background fill* (and border) of data labels on a
chart — what PowerPoint's "Format Data Labels" pane calls the *fill*
of the label's containing rectangle. At report time ``DataLabels`` and
``DataLabel`` had no ``.format`` property, so there was no clean API
for setting fill / line on label boxes.

This gap was closed by:

* Wave 1 ``feat/issue-716-datalabel-border`` (commit af2ec193) — added
  ``DataLabel.format`` (``ChartFormat`` wrapping the individual ``c:dLbl``)
  with ``.fill`` / ``.line`` / ``.shadow``.

* Wave 7 ``feat/issue-560-data-label-colors`` (commit 2018024d) — added
  ``DataLabels.format`` (``ChartFormat`` wrapping the series-level
  ``c:dLbls``) with ``.fill`` / ``.line`` / ``.shadow`` plus the
  ``c:spPr`` registration on ``CT_DLbls`` so the element is inserted
  in correct schema order.

Together, these resolve #662. This module exercises the reporter's
exact use case — "set the background fill on data labels" — at both
the per-series and per-point scope, round-tripping through save +
reload against a real package.

Leader lines (the second half of the reporter's comment thread) are
tracked separately and remain out of scope here.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.chtfmt import ChartFormat
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches


def _build_chart_with_label_backgrounds():
    """Build a single-series column chart, set label backgrounds, round-trip."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("s1", (1.0, 2.0, 3.0))
    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        chart_data,
    )
    chart = chart_shape.chart
    chart.plots[0].has_data_labels = True
    series = chart.plots[0].series[0]

    # ---#662 core: set the background fill on every label in the series---
    series.data_labels.format.fill.solid()
    series.data_labels.format.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0x00)

    # ---#662 also: per-point label background fill---
    series.points[0].data_label.format.fill.solid()
    series.points[0].data_label.format.fill.fore_color.rgb = RGBColor(0xAA, 0xBB, 0xCC)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


def _reloaded_series(prs):
    for shp in prs.slides[0].shapes:
        if shp.has_chart:
            return shp.chart.plots[0].series[0]
    raise AssertionError("reloaded presentation contains no chart")


class DescribeIssue662DataLabelBackground(object):
    """Verify #662 — data-label background fill — is resolved."""

    def it_sets_a_solid_background_on_all_series_data_labels(self):
        """Reporter's exact use case: one API call colors every label.

        ``series.data_labels.format.fill.solid()`` +
        ``fore_color.rgb = ...`` must write a valid ``c:dLbls/c:spPr``
        that survives save + reload. Before Wave 7 #560,
        ``DataLabels.format`` did not exist.
        """
        prs = _build_chart_with_label_backgrounds()
        series = _reloaded_series(prs)

        fmt = series.data_labels.format
        assert isinstance(fmt, ChartFormat)
        assert fmt.fill.fore_color.rgb == RGBColor(0xFF, 0xFF, 0x00)

    def it_sets_a_solid_background_on_an_individual_data_label(self):
        """Per-point variant: ``points[i].data_label.format.fill`` — shipped by #716."""
        prs = _build_chart_with_label_backgrounds()
        series = _reloaded_series(prs)

        fmt = series.points[0].data_label.format
        assert isinstance(fmt, ChartFormat)
        assert fmt.fill.fore_color.rgb == RGBColor(0xAA, 0xBB, 0xCC)

    def it_writes_c_spPr_under_c_dLbls_in_schema_order(self):
        """#560 specifically registered ``c:spPr`` on ``CT_DLbls``.

        Pin the in-XML placement so a future refactor that drops that
        registration would fail here instead of silently producing a
        PPTX that PowerPoint reports as needing repair.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        cd = CategoryChartData()
        cd.categories = ["A", "B"]
        cd.add_series("s1", (1.0, 2.0))
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            cd,
        ).chart
        chart.plots[0].has_data_labels = True
        series = chart.plots[0].series[0]

        series.data_labels.format.fill.solid()
        series.data_labels.format.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)

        dLbls = series.data_labels._element
        # ---c:spPr must be a direct child of c:dLbls---
        spPr_list = dLbls.findall(
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}spPr"
        )
        assert len(spPr_list) == 1
