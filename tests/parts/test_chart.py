# encoding: utf-8

"""
Test suite for pptx.parts.chart module.
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.chart import Chart
from pptx.oxml.chart.chart import CT_ChartSpace
from pptx.parts.chart import ChartPart

from ..unitutil.mock import class_mock, instance_mock


class DescribeChartPart(object):

    def it_provides_access_to_the_chart_object(self, chart_fixture):
        chart_part, chart_, Chart_ = chart_fixture
        chart = chart_part.chart
        Chart_.assert_called_once_with(chart_part._element, chart_part)
        assert chart is chart_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def chart_fixture(self, chartSpace_, Chart_, chart_):
        chart_part = ChartPart(None, None, chartSpace_)
        return chart_part, chart_, Chart_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Chart_(self, request, chart_):
        return class_mock(
            request, 'pptx.parts.chart.Chart', return_value=chart_
        )

    @pytest.fixture
    def chart_(self, request):
        return instance_mock(request, Chart)

    @pytest.fixture
    def chartSpace_(self, request):
        return instance_mock(request, CT_ChartSpace)
