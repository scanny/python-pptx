# pyright: reportPrivateUsage=false

"""Verify-close regression test for issue #617 — trendlines on column charts.

Issue #617 (https://github.com/scanny/python-pptx/issues/617) reported that
trendlines can be added to line charts but not to column charts. In the
current codebase ``add_trendline`` and ``trendlines`` live on
:class:`~pptx.chart.series._BaseSeries`, which |BarSeries| inherits via
|_BaseCategorySeries|. Column charts (``XL_CHART_TYPE.COLUMN_CLUSTERED`` and
its stacked / 3-D variants) render as ``c:barChart`` with ``c:barDir=col``,
so their series are exactly the same |BarSeries| objects that bar charts
use. Trendlines therefore work on column-chart series in exactly the same
way they work on line-chart series.

This suite pins that symmetry end-to-end — every
:class:`~pptx.enum.chart.XL_TRENDLINE_TYPE` member round-trips, survives a
``Presentation.save`` + reopen cycle, and lands in the ``c:ser`` element as
the expected ``c:trendline/c:trendlineType`` fragment — so a future
refactor that accidentally gates trendline support behind a chart-family
check would be caught here.
"""

from __future__ import annotations

import io
from typing import TYPE_CHECKING, cast

import pytest

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.chart.series import BarSeries, Trendline
from pptx.enum.chart import XL_CHART_TYPE, XL_TRENDLINE_TYPE
from pptx.shapes.graphfrm import GraphicFrame

if TYPE_CHECKING:
    from pptx.chart.chart import Chart

ALL_TRENDLINE_TYPES = [
    XL_TRENDLINE_TYPE.LINEAR,
    XL_TRENDLINE_TYPE.LOGARITHMIC,
    XL_TRENDLINE_TYPE.POLYNOMIAL,
    XL_TRENDLINE_TYPE.POWER,
    XL_TRENDLINE_TYPE.EXPONENTIAL,
    XL_TRENDLINE_TYPE.MOVING_AVG,
]


class DescribeIssue617ColumnTrendlineVerify:
    """Verify-close regression suite for trendlines on column-chart series."""

    def it_exposes_add_trendline_and_trendlines_on_column_series(self):
        # -- ``_BaseSeries.add_trendline`` / ``.trendlines`` are inherited
        # -- by |BarSeries|, which serves COLUMN_CLUSTERED charts (the
        # -- ``c:barChart`` element distinguishes column from bar via
        # -- ``c:barDir`` only — the series class is the same).
        chart = _new_column_chart()
        series = chart.plots[0].series[0]

        assert isinstance(series, BarSeries)
        assert hasattr(series, "add_trendline")
        assert hasattr(series, "trendlines")
        assert series.trendlines == []

    def it_attaches_a_linear_trendline_to_a_column_series(self):
        # -- the canonical reproducer: COLUMN_CLUSTERED + LINEAR trendline --
        chart = _new_column_chart()
        series = chart.plots[0].series[0]

        tl = series.add_trendline(XL_TRENDLINE_TYPE.LINEAR)

        assert isinstance(tl, Trendline)
        assert tl.trendline_type == XL_TRENDLINE_TYPE.LINEAR
        assert len(series.trendlines) == 1

    @pytest.mark.parametrize("trendline_type", ALL_TRENDLINE_TYPES)
    def it_accepts_every_XL_TRENDLINE_TYPE_on_a_column_series(self, trendline_type):
        # -- PR #617's brief lists linear, log, polynomial, power, exp, and
        # -- moving-average. Pin each one individually.
        chart = _new_column_chart()
        series = chart.plots[0].series[0]

        kwargs = {}
        if trendline_type == XL_TRENDLINE_TYPE.POLYNOMIAL:
            kwargs["order"] = 3
        elif trendline_type == XL_TRENDLINE_TYPE.MOVING_AVG:
            kwargs["period"] = 2

        tl = series.add_trendline(trendline_type, **kwargs)

        assert tl.trendline_type == trendline_type

    def it_attaches_multiple_trendlines_to_the_same_column_series(self):
        # -- one series, three distinct fits — the schema allows any number.
        chart = _new_column_chart()
        series = chart.plots[0].series[0]

        series.add_trendline(XL_TRENDLINE_TYPE.LINEAR)
        series.add_trendline(XL_TRENDLINE_TYPE.POLYNOMIAL, order=2)
        series.add_trendline(XL_TRENDLINE_TYPE.MOVING_AVG, period=3)

        types = [tl.trendline_type for tl in series.trendlines]
        assert types == [
            XL_TRENDLINE_TYPE.LINEAR,
            XL_TRENDLINE_TYPE.POLYNOMIAL,
            XL_TRENDLINE_TYPE.MOVING_AVG,
        ]

    def it_honours_display_equation_and_r_squared_on_a_column_trendline(self):
        # -- the on-chart equation and R-squared labels are part of the
        # -- advertised feature surface.
        chart = _new_column_chart()
        series = chart.plots[0].series[0]

        tl = series.add_trendline(
            XL_TRENDLINE_TYPE.LINEAR,
            display_equation=True,
            display_r_squared=True,
        )

        assert tl.display_equation is True
        assert tl.display_r_squared is True

    def it_survives_a_save_reopen_roundtrip(self):
        # -- trendlines are a property of authored XML, so the full
        # -- complement must round-trip through
        # -- ``Presentation.save`` + reopen with no drift.
        prs, chart = _new_column_chart_with_prs()
        series = chart.plots[0].series[0]
        for tl_type in ALL_TRENDLINE_TYPES:
            kwargs = {}
            if tl_type == XL_TRENDLINE_TYPE.POLYNOMIAL:
                kwargs["order"] = 3
            elif tl_type == XL_TRENDLINE_TYPE.MOVING_AVG:
                kwargs["period"] = 2
            series.add_trendline(tl_type, **kwargs)

        reloaded = _roundtrip(prs)
        gf = cast(
            GraphicFrame,
            next(s for s in reloaded.slides[0].shapes if s.has_chart),
        )
        reloaded_series = gf.chart.plots[0].series[0]

        assert isinstance(reloaded_series, BarSeries)
        assert [tl.trendline_type for tl in reloaded_series.trendlines] == ALL_TRENDLINE_TYPES

    def it_writes_a_c_trendline_child_under_the_c_ser_in_a_c_barChart(self):
        # -- pin the XML surface: a ``c:trendline`` child must appear under
        # -- the ``c:ser`` element, which is itself a child of ``c:barChart``
        # -- (column charts render as ``c:barChart`` + ``c:barDir=col``).
        chart = _new_column_chart()
        series = chart.plots[0].series[0]
        series.add_trendline(XL_TRENDLINE_TYPE.LINEAR)

        ser = series._ser
        trendlines = ser.findall(_qn("c:trendline"))
        assert len(trendlines) == 1
        trendline_types = trendlines[0].findall(_qn("c:trendlineType"))
        assert len(trendline_types) == 1
        assert trendline_types[0].get("val") == "linear"
        # -- the parent xChart tag is c:barChart, not c:lineChart --
        assert ser.getparent().tag == _qn("c:barChart")

    def it_also_works_on_every_other_column_chart_variant(self):
        # -- COLUMN_CLUSTERED is the canonical repro, but the same BarSeries
        # -- powers every other column variant. Pin one trendline per
        # -- variant to catch any future chart-type gating.
        variants = [
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            XL_CHART_TYPE.COLUMN_STACKED,
            XL_CHART_TYPE.COLUMN_STACKED_100,
        ]
        for variant in variants:
            chart = _new_chart(variant)
            series = chart.plots[0].series[0]
            tl = series.add_trendline(XL_TRENDLINE_TYPE.LINEAR)
            assert tl.trendline_type == XL_TRENDLINE_TYPE.LINEAR, variant
            assert len(series.trendlines) == 1, variant


# -- helpers --------------------------------------------------------------


def _new_column_chart() -> "Chart":
    """Return a chart on a fresh slide in a fresh presentation."""
    _, chart = _new_column_chart_with_prs()
    return chart


def _new_column_chart_with_prs():
    """Return ``(presentation, chart)`` for a fresh COLUMN_CLUSTERED chart."""
    return _new_chart_with_prs(XL_CHART_TYPE.COLUMN_CLUSTERED)


def _new_chart(chart_type: XL_CHART_TYPE) -> "Chart":
    _, chart = _new_chart_with_prs(chart_type)
    return chart


def _new_chart_with_prs(chart_type: XL_CHART_TYPE):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    data = CategoryChartData()
    data.categories = ["Q1", "Q2", "Q3", "Q4"]
    data.add_series("S", (1.0, 2.0, 3.0, 4.5))
    gf = slide.shapes.add_chart(chart_type, 0, 0, 0, 0, data)
    return prs, gf.chart


def _roundtrip(prs):
    """Serialize ``prs`` to a ``BytesIO`` and reopen it."""
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


def _qn(tag):
    """Return a Clark-notation tag name for the given ``c:``-prefixed tag."""
    from pptx.oxml.ns import qn

    return qn(tag)
