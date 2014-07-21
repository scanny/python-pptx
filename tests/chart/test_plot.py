# encoding: utf-8

"""
Test suite for pptx.chart.plot module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.plot import (
    BarPlot, DataLabels, LinePlot, PiePlot, Plot, PlotFactory
)

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class DescribePlot(object):

    def it_knows_whether_it_has_data_labels(
            self, has_data_labels_get_fixture):
        plot, expected_value = has_data_labels_get_fixture
        assert plot.has_data_labels == expected_value

    def it_can_change_whether_it_has_data_labels(
            self, has_data_labels_set_fixture):
        plot, new_value, expected_xml = has_data_labels_set_fixture
        plot.has_data_labels = new_value
        assert plot._element.xml == expected_xml

    def it_provides_access_to_the_data_labels(self, data_labels_fixture):
        plot, data_labels_, DataLabels_, dLbls = data_labels_fixture
        data_labels = plot.data_labels
        DataLabels_.assert_called_once_with(dLbls)
        assert data_labels is data_labels_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def data_labels_fixture(self, DataLabels_, data_labels_):
        barChart = element('c:barChart/c:dLbls')
        dLbls = barChart[0]
        plot = Plot(barChart)
        return plot, data_labels_, DataLabels_, dLbls

    @pytest.fixture(params=[
        ('c:barChart',  False), ('c:barChart/c:dLbls',  True),
        ('c:lineChart', False), ('c:lineChart/c:dLbls', True),
        ('c:pieChart',  False), ('c:pieChart/c:dLbls',  True),
    ])
    def has_data_labels_get_fixture(self, request):
        plot_cxml, expected_value = request.param
        plot = Plot(element(plot_cxml))
        return plot, expected_value

    @pytest.fixture(params=[
        ('c:barChart',          True,  'c:barChart/c:dLbls/+'),
        ('c:barChart/c:dLbls',  True,  'c:barChart/c:dLbls'),
        ('c:barChart',          False, 'c:barChart'),
        ('c:barChart/c:dLbls',  False, 'c:barChart'),
        ('c:lineChart',         True,  'c:lineChart/c:dLbls/+'),
        ('c:lineChart/c:dLbls', False, 'c:lineChart'),
        ('c:pieChart',          True,  'c:pieChart/c:dLbls/+'),
        ('c:pieChart/c:dLbls',  False, 'c:pieChart'),
    ])
    def has_data_labels_set_fixture(self, request):
        plot_cxml, new_value, expected_plot_cxml = request.param
        # apply extended suffix to replace trailing '+' where present
        if expected_plot_cxml.endswith('+'):
            expected_plot_cxml = expected_plot_cxml[:-1] + (
                '(c:showLegendKey{val=0},c:showVal{val=1},c:showCatName{val='
                '0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSi'
                'ze{val=0})'
            )
        plot = PlotFactory(element(plot_cxml))
        expected_xml = xml(expected_plot_cxml)
        return plot, new_value, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def DataLabels_(self, request, data_labels_):
        return class_mock(
            request, 'pptx.chart.plot.DataLabels',
            return_value=data_labels_
        )

    @pytest.fixture
    def data_labels_(self, request):
        return instance_mock(request, DataLabels)


class DescribeDataLabels(object):

    def it_knows_its_number_format(self, number_format_get_fixture):
        data_labels, expected_value = number_format_get_fixture
        assert data_labels.number_format == expected_value

    def it_can_change_its_number_format(self, number_format_set_fixture):
        data_labels, new_value, expected_xml = number_format_set_fixture
        data_labels.number_format = new_value
        assert data_labels._element.xml == expected_xml

    def it_knows_whether_its_number_format_is_linked(
            self, number_format_is_linked_get_fixture):
        data_labels, expected_value = number_format_is_linked_get_fixture
        assert data_labels.number_format_is_linked is expected_value

    def it_can_change_whether_its_number_format_is_linked(
            self, number_format_is_linked_set_fixture):
        data_labels, new_value, expected_xml = (
            number_format_is_linked_set_fixture
        )
        data_labels.number_format_is_linked = new_value
        assert data_labels._element.xml == expected_xml

    @pytest.fixture(params=[
        ('c:dLbls', True,  'c:dLbls/c:numFmt{sourceLinked=1}'),
        ('c:dLbls', False, 'c:dLbls/c:numFmt{sourceLinked=0}'),
        ('c:dLbls', None,  'c:dLbls/c:numFmt'),
        ('c:dLbls/c:numFmt', True, 'c:dLbls/c:numFmt{sourceLinked=1}'),
        ('c:dLbls/c:numFmt{sourceLinked=1}', False,
         'c:dLbls/c:numFmt{sourceLinked=0}'),
    ])
    def number_format_is_linked_set_fixture(self, request):
        dLbls_cxml, new_value, expected_dLbls_cxml = request.param
        data_labels = DataLabels(element(dLbls_cxml))
        expected_xml = xml(expected_dLbls_cxml)
        return data_labels, new_value, expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:dLbls',                             'General'),
        ('c:dLbls/c:numFmt{formatCode=foobar}', 'foobar'),
    ])
    def number_format_get_fixture(self, request):
        dLbls_cxml, expected_value = request.param
        data_labels = DataLabels(element(dLbls_cxml))
        return data_labels, expected_value

    @pytest.fixture(params=[
        ('c:dLbls', 'General',
         'c:dLbls/c:numFmt{formatCode=General,sourceLinked=0}'),
        ('c:dLbls/c:numFmt{formatCode=General}', '00.00',
         'c:dLbls/c:numFmt{formatCode=00.00,sourceLinked=0}'),
    ])
    def number_format_set_fixture(self, request):
        dLbls_cxml, new_value, expected_dLbls_cxml = request.param
        data_labels = DataLabels(element(dLbls_cxml))
        expected_xml = xml(expected_dLbls_cxml)
        return data_labels, new_value, expected_xml

    @pytest.fixture(params=[
        ('c:dLbls',                          True),
        ('c:dLbls/c:numFmt',                 True),
        ('c:dLbls/c:numFmt{sourceLinked=0}', False),
        ('c:dLbls/c:numFmt{sourceLinked=1}', True),
    ])
    def number_format_is_linked_get_fixture(self, request):
        dLbls_cxml, expected_value = request.param
        data_labels = DataLabels(element(dLbls_cxml))
        return data_labels, expected_value


class DescribePlotFactory(object):

    def it_contructs_a_plot_object_from_a_plot_element(self, call_fixture):
        plot_elm, PlotClass_, plot_ = call_fixture
        plot = PlotFactory(plot_elm)
        PlotClass_.assert_called_once_with(plot_elm)
        assert plot is plot_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        'barChart',
        'lineChart',
        'pieChart'
    ])
    def call_fixture(
            self, request, BarPlot_, bar_chart_, LinePlot_, line_chart_,
            PiePlot_, pie_chart_):
        plot_cxml, PlotClass_, plot_mock = {
            'barChart':  ('c:barChart',  BarPlot_,  bar_chart_),
            'lineChart': ('c:lineChart', LinePlot_, line_chart_),
            'pieChart':  ('c:pieChart',  PiePlot_,  pie_chart_),
        }[request.param]
        plot_elm = element(plot_cxml)
        return plot_elm, PlotClass_, plot_mock

    # fixture components -----------------------------------

    @pytest.fixture
    def BarPlot_(self, request, bar_chart_):
        return class_mock(
            request, 'pptx.chart.plot.BarPlot',
            return_value=bar_chart_
        )

    @pytest.fixture
    def bar_chart_(self, request):
        return instance_mock(request, BarPlot)

    @pytest.fixture
    def LinePlot_(self, request, line_chart_):
        return class_mock(
            request, 'pptx.chart.plot.LinePlot',
            return_value=line_chart_
        )

    @pytest.fixture
    def line_chart_(self, request):
        return instance_mock(request, LinePlot)

    @pytest.fixture
    def PiePlot_(self, request, pie_chart_):
        return class_mock(
            request, 'pptx.chart.plot.PiePlot',
            return_value=pie_chart_
        )

    @pytest.fixture
    def pie_chart_(self, request):
        return instance_mock(request, PiePlot)
