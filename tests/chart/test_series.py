# encoding: utf-8

"""
Test suite for pptx.chart.series module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.series import BarSeries, LineSeries, PieSeries, SeriesFactory

from ..unitutil.cxml import element
from ..unitutil.mock import class_mock, instance_mock


class DescribeSeriesFactory(object):

    def it_contructs_a_series_object_from_a_plot_element(self, call_fixture):
        xChart, ser, SeriesCls_, series_ = call_fixture
        series = SeriesFactory(xChart, ser)
        SeriesCls_.assert_called_once_with(ser)
        assert series is series_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        'barChart',
        'lineChart',
        'pieChart'
    ])
    def call_fixture(
            self, request, BarSeries_, bar_series_, LineSeries_,
            line_series_, PieSeries_, pie_series_):
        xChart_cxml, SeriesCls_, series_ = {
            'barChart':  ('c:barChart/c:ser',  BarSeries_,  bar_series_),
            'lineChart': ('c:lineChart/c:ser', LineSeries_, line_series_),
            'pieChart':  ('c:pieChart/c:ser',  PieSeries_,  pie_series_),
        }[request.param]
        xChart = element(xChart_cxml)
        ser = xChart[0]
        return xChart, ser, SeriesCls_, series_

    # fixture components -----------------------------------

    @pytest.fixture
    def BarSeries_(self, request, bar_series_):
        return class_mock(
            request, 'pptx.chart.series.BarSeries',
            return_value=bar_series_
        )

    @pytest.fixture
    def bar_series_(self, request):
        return instance_mock(request, BarSeries)

    @pytest.fixture
    def LineSeries_(self, request, line_series_):
        return class_mock(
            request, 'pptx.chart.series.LineSeries',
            return_value=line_series_
        )

    @pytest.fixture
    def line_series_(self, request):
        return instance_mock(request, LineSeries)

    @pytest.fixture
    def PieSeries_(self, request, pie_series_):
        return class_mock(
            request, 'pptx.chart.series.PieSeries',
            return_value=pie_series_
        )

    @pytest.fixture
    def pie_series_(self, request):
        return instance_mock(request, PieSeries)
