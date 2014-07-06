# encoding: utf-8

"""
Test suite for pptx.chart module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.axis import CategoryAxis, ValueAxis
from pptx.chart.chart import Chart, Plots

from ..unitutil.cxml import element, xml
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

    def it_provides_access_to_the_value_axis(self, val_ax_fixture):
        chart, value_axis_, ValueAxis_, valAx = val_ax_fixture
        value_axis = chart.value_axis
        ValueAxis_.assert_called_once_with(valAx)
        assert value_axis is value_axis_

    def it_raises_when_no_value_axis(self, val_ax_raise_fixture):
        chart = val_ax_raise_fixture
        with pytest.raises(ValueError):
            chart.value_axis

    def it_provides_access_to_the_plots(self, plots_fixture):
        chart, plots_, Plots_, plotArea = plots_fixture
        plots = chart.plots
        Plots_.assert_called_once_with(plotArea)
        assert plots is plots_

    def it_knows_its_style(self, style_get_fixture):
        chart, expected_value = style_get_fixture
        assert chart.chart_style == expected_value

    def it_can_change_its_style(self, style_set_fixture):
        chart, new_value, expected_xml = style_set_fixture
        chart.chart_style = new_value
        assert chart._chartSpace.xml == expected_xml

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

    @pytest.fixture
    def plots_fixture(self, Plots_, plots_):
        chartSpace = element('c:chartSpace/c:chart/c:plotArea')
        plotArea = chartSpace.xpath('./c:chart/c:plotArea')[0]
        chart = Chart(chartSpace, None)
        return chart, plots_, Plots_, plotArea

    @pytest.fixture(params=[
        ('c:chartSpace/c:style{val=42}', 42),
        ('c:chartSpace',                 None),
    ])
    def style_get_fixture(self, request):
        chartSpace_cxml, expected_value = request.param
        chart = Chart(element(chartSpace_cxml), None)
        return chart, expected_value

    @pytest.fixture(params=[
        ('c:chartSpace',                4,    'c:chartSpace/c:style{val=4}'),
        ('c:chartSpace',                None, 'c:chartSpace'),
        ('c:chartSpace/c:style{val=4}', 2,    'c:chartSpace/c:style{val=2}'),
        ('c:chartSpace/c:style{val=4}', None, 'c:chartSpace'),
    ])
    def style_set_fixture(self, request):
        chartSpace_cxml, new_value, expected_chartSpace_cxml = request.param
        chart = Chart(element(chartSpace_cxml), None)
        expected_xml = xml(expected_chartSpace_cxml)
        return chart, new_value, expected_xml

    @pytest.fixture
    def val_ax_fixture(self, ValueAxis_, value_axis_):
        chartSpace = element('c:chartSpace/c:chart/c:plotArea/c:valAx')
        valAx = chartSpace.xpath('./c:chart/c:plotArea/c:valAx')[0]
        chart = Chart(chartSpace, None)
        return chart, value_axis_, ValueAxis_, valAx

    @pytest.fixture
    def val_ax_raise_fixture(self):
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

    @pytest.fixture
    def Plots_(self, request, plots_):
        return class_mock(
            request, 'pptx.chart.chart.Plots', return_value=plots_
        )

    @pytest.fixture
    def plots_(self, request):
        return instance_mock(request, Plots)

    @pytest.fixture
    def ValueAxis_(self, request, value_axis_):
        return class_mock(
            request, 'pptx.chart.chart.ValueAxis',
            return_value=value_axis_
        )

    @pytest.fixture
    def value_axis_(self, request):
        return instance_mock(request, ValueAxis)
