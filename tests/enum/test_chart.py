"""Unit-test suite for `pptx.enum.chart` module."""

from __future__ import annotations

from pptx.enum.chart import XL_CHART_TYPE


class DescribeXL_CHART_TYPE(object):
    """Unit-test suite for `pptx.enum.chart.XL_CHART_TYPE` enumeration."""

    def it_exposes_a_sentinel_for_unsupported_chartex_charts(self):
        assert hasattr(XL_CHART_TYPE, "UNSUPPORTED_CHARTEX")
        assert XL_CHART_TYPE.UNSUPPORTED_CHARTEX in XL_CHART_TYPE

    def and_the_sentinel_value_is_outside_the_Microsoft_API_range(self):
        # -- Microsoft API `XlChartType` values fall in [-4169, 112]; we use -9999 so the
        # -- value is unambiguously a python-pptx sentinel and cannot collide with a future
        # -- Microsoft API addition that lands inside the legitimate range.
        assert XL_CHART_TYPE.UNSUPPORTED_CHARTEX.value == -9999

    def and_the_sentinel_renders_its_name_in_str(self):
        assert "UNSUPPORTED_CHARTEX" in str(XL_CHART_TYPE.UNSUPPORTED_CHARTEX)
