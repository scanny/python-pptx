# encoding: utf-8

"""
Test suite for pptx.chart.chart module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.axis import CategoryAxis, ValueAxis
from pptx.chart.chart import Chart, Legend, _Plots, _SeriesRewriter
from pptx.chart.data import ChartData, _SeriesData
from pptx.chart.plot import Plot
from pptx.chart.series import SeriesCollection
from pptx.enum.base import EnumValue
from pptx.oxml import parse_xml
from pptx.oxml.chart.chart import CT_ChartSpace
from pptx.oxml.chart.series import CT_SeriesComposite
from pptx.parts.chart import ChartPart

from ..unitutil.cxml import element, xml
from ..unitutil.file import snippet_seq
from ..unitutil.mock import (
    call, class_mock, function_mock, instance_mock, method_mock
)


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

    def it_provides_access_to_its_series(self, series_fixture):
        chart, SeriesCollection_, chartSpace_, series_ = series_fixture
        series = chart.series
        SeriesCollection_.assert_called_once_with(chartSpace_)
        assert series is series_

    def it_provides_access_to_its_plots(self, plots_fixture):
        chart, plots_, _Plots_, plotArea = plots_fixture
        plots = chart.plots
        _Plots_.assert_called_once_with(plotArea, chart)
        assert plots is plots_

    def it_knows_whether_it_has_a_legend(self, has_legend_get_fixture):
        chart, expected_value = has_legend_get_fixture
        assert chart.has_legend == expected_value

    def it_can_change_whether_it_has_a_legend(self, has_legend_set_fixture):
        chart, new_value, expected_xml = has_legend_set_fixture
        chart.has_legend = new_value
        assert chart._chartSpace.xml == expected_xml

    def it_provides_access_to_its_legend(self, legend_fixture):
        chart, Legend_, expected_calls, expected_value = legend_fixture
        legend = chart.legend
        assert Legend_.call_args_list == expected_calls
        assert legend is expected_value

    def it_knows_its_chart_type(self, chart_type_fixture):
        chart, PlotTypeInspector_, plot_, chart_type_ = chart_type_fixture
        chart_type = chart.chart_type
        PlotTypeInspector_.chart_type.assert_called_once_with(plot_)
        assert chart_type is chart_type_

    def it_knows_its_style(self, style_get_fixture):
        chart, expected_value = style_get_fixture
        assert chart.chart_style == expected_value

    def it_can_change_its_style(self, style_set_fixture):
        chart, new_value, expected_xml = style_set_fixture
        chart.chart_style = new_value
        assert chart._chartSpace.xml == expected_xml

    def it_can_replace_the_chart_data(self, replace_data_fixture):
        chart, chart_data_, _SeriesRewriter_, chartSpace_, workbook_ = (
            replace_data_fixture
        )
        chart.replace_data(chart_data_)
        _SeriesRewriter_.replace_series_data.assert_called_once_with(
            chartSpace_, chart_data_
        )
        workbook_.update_from_xlsx_blob.assert_called_once_with(
            chart_data_.xlsx_blob
        )

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
    def chart_type_fixture(self, PlotTypeInspector_, plot_, chart_type_):
        chart = Chart(None, None)
        chart._plots = [plot_]
        return chart, PlotTypeInspector_, plot_, chart_type_

    @pytest.fixture(params=[
        ('c:chartSpace/c:chart',          False),
        ('c:chartSpace/c:chart/c:legend', True),
    ])
    def has_legend_get_fixture(self, request):
        chartSpace_cxml, expected_value = request.param
        chart = Chart(element(chartSpace_cxml), None)
        return chart, expected_value

    @pytest.fixture(params=[
        ('c:chartSpace/c:chart', True,
         'c:chartSpace/c:chart/c:legend'),
    ])
    def has_legend_set_fixture(self, request):
        chartSpace_cxml, new_value, expected_chartSpace_cxml = request.param
        chart = Chart(element(chartSpace_cxml), None)
        expected_xml = xml(expected_chartSpace_cxml)
        return chart, new_value, expected_xml

    @pytest.fixture(params=[
        ('c:chartSpace/c:chart',          False),
        ('c:chartSpace/c:chart/c:legend', True),
    ])
    def legend_fixture(self, request, Legend_, legend_):
        chartSpace_cxml, has_legend = request.param
        chartSpace = element(chartSpace_cxml)
        chart = Chart(chartSpace, None)
        expected_value, expected_calls = None, []
        if has_legend:
            expected_value = legend_
            legend_elm = chartSpace.chart.legend
            expected_calls.append(call(legend_elm))
        return chart, Legend_, expected_calls, expected_value

    @pytest.fixture
    def plots_fixture(self, _Plots_, plots_):
        chartSpace = element('c:chartSpace/c:chart/c:plotArea')
        plotArea = chartSpace.xpath('./c:chart/c:plotArea')[0]
        chart = Chart(chartSpace, None)
        return chart, plots_, _Plots_, plotArea

    @pytest.fixture
    def replace_data_fixture(
            self, chartSpace_, chart_part_, chart_data_, _SeriesRewriter_):
        chart = Chart(chartSpace_, chart_part_)
        workbook_ = chart_part_.chart_workbook
        return (
            chart, chart_data_, _SeriesRewriter_, chartSpace_, workbook_
        )

    @pytest.fixture
    def series_fixture(
            self, SeriesCollection_, chartSpace_, series_collection_):
        chart = Chart(chartSpace_, None)
        return chart, SeriesCollection_, chartSpace_, series_collection_

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
    def chartSpace_(self, request):
        return instance_mock(request, CT_ChartSpace)

    @pytest.fixture
    def chart_data_(self, request):
        return instance_mock(request, ChartData)

    @pytest.fixture
    def chart_part_(self, request):
        return instance_mock(request, ChartPart)

    @pytest.fixture
    def chart_type_(self, request):
        return instance_mock(request, EnumValue)

    @pytest.fixture
    def Legend_(self, request, legend_):
        return class_mock(
            request, 'pptx.chart.chart.Legend', return_value=legend_
        )

    @pytest.fixture
    def legend_(self, request):
        return instance_mock(request, Legend)

    @pytest.fixture
    def PlotTypeInspector_(self, request, chart_type_):
        PlotTypeInspector_ = class_mock(
            request, 'pptx.chart.chart.PlotTypeInspector'
        )
        PlotTypeInspector_.chart_type.return_value = chart_type_
        return PlotTypeInspector_

    @pytest.fixture
    def _Plots_(self, request, plots_):
        return class_mock(
            request, 'pptx.chart.chart._Plots', return_value=plots_
        )

    @pytest.fixture
    def plot_(self, request):
        return instance_mock(request, Plot)

    @pytest.fixture
    def plots_(self, request):
        return instance_mock(request, _Plots)

    @pytest.fixture
    def SeriesCollection_(self, request, series_collection_):
        return class_mock(
            request, 'pptx.chart.chart.SeriesCollection',
            return_value=series_collection_
        )

    @pytest.fixture
    def _SeriesRewriter_(self, request):
        return class_mock(request, 'pptx.chart.chart._SeriesRewriter')

    @pytest.fixture
    def series_collection_(self, request):
        return instance_mock(request, SeriesCollection)

    @pytest.fixture
    def ValueAxis_(self, request, value_axis_):
        return class_mock(
            request, 'pptx.chart.chart.ValueAxis',
            return_value=value_axis_
        )

    @pytest.fixture
    def value_axis_(self, request):
        return instance_mock(request, ValueAxis)


class Describe_Plots(object):

    def it_supports_indexed_access(self, getitem_fixture):
        plots, idx, PlotFactory_, plot_elm, chart_, plot_ = getitem_fixture
        plot = plots[idx]
        PlotFactory_.assert_called_once_with(plot_elm, chart_)
        assert plot is plot_

    def it_supports_len(self, len_fixture):
        plots, expected_len = len_fixture
        assert len(plots) == expected_len

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:plotArea/c:barChart', 0),
        ('c:plotArea/(c:radarChart,c:barChart)', 1),
    ])
    def getitem_fixture(self, request, PlotFactory_, chart_, plot_):
        plotArea_cxml, idx = request.param
        plotArea = element(plotArea_cxml)
        plot_elm = plotArea[idx]
        plots = _Plots(plotArea, chart_)
        return plots, idx, PlotFactory_, plot_elm, chart_, plot_

    @pytest.fixture(params=[
        ('c:plotArea',                          0),
        ('c:plotArea/c:barChart',               1),
        ('c:plotArea/(c:barChart,c:lineChart)', 2),
    ])
    def len_fixture(self, request):
        plotArea_cxml, expected_len = request.param
        plots = _Plots(element(plotArea_cxml), None)
        return plots, expected_len

    # fixture components ---------------------------------------------

    @pytest.fixture
    def chart_(self, request):
        return instance_mock(request, Chart)

    @pytest.fixture
    def PlotFactory_(self, request, plot_):
        return function_mock(
            request, 'pptx.chart.chart.PlotFactory', return_value=plot_
        )

    @pytest.fixture
    def plot_(self, request):
        return instance_mock(request, Plot)


class Describe_SeriesRewriter(object):

    def it_can_replace_the_sers_in_a_chartSpace(self, replace_fixture):
        chartSpace_, chart_data_, series_count, expected_calls = (
            replace_fixture
        )

        _SeriesRewriter.replace_series_data(chartSpace_, chart_data_)

        _SeriesRewriter._adjust_ser_count.assert_called_once_with(
            chartSpace_, series_count
        )
        assert _SeriesRewriter._rewrite_ser_data.call_args_list == (
            expected_calls
        )

    def it_can_change_the_number_of_sers_in_a_chartSpace(
            self, adjust_fixture):
        chartSpace, new_ser_count, expected_xml = adjust_fixture
        _SeriesRewriter._adjust_ser_count(chartSpace, new_ser_count)
        assert chartSpace.xml == expected_xml

    def it_rewrites_ser_data_to_help_replace_series_data(
            self, rewrite_ser_fixture):
        ser, series_data, expected_xml = rewrite_ser_fixture
        _SeriesRewriter._rewrite_ser_data(ser, series_data)
        assert ser.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        # same ser_count doesn't change anything
        (0, 2),
        # greater ser_count adds to end of last xChart
        (2, 3),
        # lesser ser_count removes from last in idx order, not doc order
        (4, 2),
        # an xChart left with no sers is removed
        (6, 1),
    ])
    def adjust_fixture(self, request):
        snippet_offset, new_ser_count = request.param
        chartSpace_xml, expected_xml = snippet_seq(
            'adjust-ser-count', snippet_offset, count=2
        )
        chartSpace = parse_xml(chartSpace_xml)
        return chartSpace, new_ser_count, expected_xml

    @pytest.fixture
    def replace_fixture(
            self, chartSpace_, chart_data_, _adjust_ser_count_,
            _rewrite_ser_data_, ser_, ser_2_, series_data_, series_data_2_):
        _adjust_ser_count_.return_value = [ser_, ser_2_]
        series_count = 2
        expected_calls = [
            call(ser_,   series_data_),
            call(ser_2_, series_data_2_)
        ]
        return chartSpace_, chart_data_, series_count, expected_calls

    @pytest.fixture
    def rewrite_ser_fixture(self):
        ser_xml, expected_xml = snippet_seq('rewrite-ser')
        ser = parse_xml(ser_xml)
        series_data = _SeriesData(2, 'Series 4', (1, 2), ('Foo', 'Bar'), 0)
        return ser, series_data, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _adjust_ser_count_(self, request):
        return method_mock(request, _SeriesRewriter, '_adjust_ser_count')

    @pytest.fixture
    def chartSpace_(self, request):
        return instance_mock(request, CT_ChartSpace)

    @pytest.fixture
    def chart_data_(self, request, series_data_, series_data_2_):
        chart_data_ = instance_mock(request, ChartData)
        chart_data_.series = [series_data_, series_data_2_]
        return chart_data_

    @pytest.fixture
    def _rewrite_ser_data_(self, request):
        return method_mock(request, _SeriesRewriter, '_rewrite_ser_data')

    @pytest.fixture
    def ser_(self, request):
        return instance_mock(request, CT_SeriesComposite)

    @pytest.fixture
    def ser_2_(self, request):
        return instance_mock(request, CT_SeriesComposite)

    @pytest.fixture
    def series_data_(self, request):
        return instance_mock(request, _SeriesData)

    @pytest.fixture
    def series_data_2_(self, request):
        return instance_mock(request, _SeriesData)
