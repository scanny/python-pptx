# encoding: utf-8

"""
Test suite for pptx.chart module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.axis import CategoryAxis
from pptx.chart.chart import Chart

from ..unitutil.cxml import element
from ..unitutil.mock import class_mock, instance_mock


class DescribeChart(object):

    def it_provides_access_to_the_category_axis(self, cat_ax_fixture):
        chart, category_axis_, CategoryAxis_, catAx = cat_ax_fixture
        category_axis = chart.category_axis
        CategoryAxis_.assert_called_once_with(catAx)
        assert category_axis is category_axis_

    def it_raises_when_no_category_axis(self, cat_ax_raise_fixture):
        chart = cat_ax_raise_fixture
        with pytest.raises(ValueError):
            chart.category_axis

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def cat_ax_fixture(self, CategoryAxis_, category_axis_):
        chartSpace = element('c:chartSpace/c:chart/c:plotArea/c:catAx')
        catAx = chartSpace.xpath('./c:chart/c:plotArea/c:catAx')[0]
        chart = Chart(chartSpace, None)
        return chart, category_axis_, CategoryAxis_, catAx

    @pytest.fixture
    def cat_ax_raise_fixture(self):
        chart = Chart(element('c:chartSpace/c:chart/c:plotArea'), None)
        return chart

    # fixture components ---------------------------------------------

    @pytest.fixture
    def CategoryAxis_(self, request, category_axis_):
        return class_mock(
            request, 'pptx.chart.chart.CategoryAxis',
            return_value=category_axis_
        )

    @pytest.fixture
    def category_axis_(self, request):
        return instance_mock(request, CategoryAxis)
