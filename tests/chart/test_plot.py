# encoding: utf-8

"""
Test suite for pptx.chart.plot module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.chart import Chart
from pptx.chart.plot import (
    AreaPlot, Area3DPlot, BarPlot, DataLabels, LinePlot, PiePlot, Plot,
    PlotFactory, PlotTypeInspector
)
from pptx.chart.series import SeriesCollection
from pptx.enum.chart import XL_CHART_TYPE as XL, XL_LABEL_POSITION
from pptx.text.text import Font

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class DescribePlot(object):

    def it_knows_which_chart_it_belongs_to(self, chart_fixture):
        plot, expected_value = chart_fixture
        assert plot.chart == expected_value

    def it_knows_its_categories(self, categories_get_fixture):
        plot, expected_value = categories_get_fixture
        assert plot.categories == expected_value

    def it_knows_whether_it_has_data_labels(
            self, has_data_labels_get_fixture):
        plot, expected_value = has_data_labels_get_fixture
        assert plot.has_data_labels == expected_value

    def it_can_change_whether_it_has_data_labels(
            self, has_data_labels_set_fixture):
        plot, new_value, expected_xml = has_data_labels_set_fixture
        plot.has_data_labels = new_value
        assert plot._element.xml == expected_xml

    def it_knows_whether_it_varies_color_by_category(
            self, vary_by_categories_get_fixture):
        plot, expected_value = vary_by_categories_get_fixture
        assert plot.vary_by_categories == expected_value

    def it_can_change_whether_it_varies_color_by_category(
            self, vary_by_categories_set_fixture):
        plot, new_value, expected_xml = vary_by_categories_set_fixture
        plot.vary_by_categories = new_value
        assert plot._element.xml == expected_xml

    def it_provides_access_to_the_data_labels(self, data_labels_fixture):
        plot, data_labels_, DataLabels_, dLbls = data_labels_fixture
        data_labels = plot.data_labels
        DataLabels_.assert_called_once_with(dLbls)
        assert data_labels is data_labels_

    def it_provides_access_to_its_series(self, series_fixture):
        plot, series_, SeriesCollection_, xChart = series_fixture
        series = plot.series
        SeriesCollection_.assert_called_once_with(xChart)
        assert series is series_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:barChart', ()),
        ('c:barChart/c:ser', ()),
        ('c:barChart/c:ser/c:cat/c:strRef', ()),
        ('c:barChart/c:ser/c:cat/c:strRef/c:strCache', ()),
        ('c:barChart/c:ser/c:cat/c:strRef/c:strCache/(c:pt{idx=1}/c:v"bar",c'
         ':pt{idx=0}/c:v"foo",c:pt{idx=2}/c:v"baz")',
         ('foo', 'bar', 'baz')),
        ('c:barChart/c:ser/c:cat/c:strLit', ()),
        ('c:barChart/c:ser/c:cat/c:strLit/(c:pt{idx=2}/c:v"faz",c:pt{idx=0}/'
         'c:v"boo",c:pt{idx=1}/c:v"far")',
         ('boo', 'far', 'faz')),
    ])
    def categories_get_fixture(self, request):
        xChart_cxml, expected_value = request.param
        plot = PlotFactory(element(xChart_cxml), None)
        return plot, expected_value

    @pytest.fixture
    def chart_fixture(self, chart_):
        plot = Plot(None, chart_)
        expected_value = chart_
        return plot, expected_value

    @pytest.fixture
    def data_labels_fixture(self, DataLabels_, data_labels_):
        barChart = element('c:barChart/c:dLbls')
        dLbls = barChart[0]
        plot = Plot(barChart, None)
        return plot, data_labels_, DataLabels_, dLbls

    @pytest.fixture(params=[
        ('c:barChart',  False), ('c:barChart/c:dLbls',  True),
        ('c:lineChart', False), ('c:lineChart/c:dLbls', True),
        ('c:pieChart',  False), ('c:pieChart/c:dLbls',  True),
    ])
    def has_data_labels_get_fixture(self, request):
        xChart_cxml, expected_value = request.param
        plot = Plot(element(xChart_cxml), None)
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
        xChart_cxml, new_value, expected_xChart_cxml = request.param
        # apply extended suffix to replace trailing '+' where present
        if expected_xChart_cxml.endswith('+'):
            expected_xChart_cxml = expected_xChart_cxml[:-1] + (
                '(c:showLegendKey{val=0},c:showVal{val=1},c:showCatName{val='
                '0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSi'
                'ze{val=0})'
            )
        plot = PlotFactory(element(xChart_cxml), None)
        expected_xml = xml(expected_xChart_cxml)
        return plot, new_value, expected_xml

    @pytest.fixture
    def series_fixture(self, SeriesCollection_, series_collection_):
        xChart = element('c:barChart')
        plot = Plot(xChart, None)
        return plot, series_collection_, SeriesCollection_, xChart

    @pytest.fixture(params=[
        ('c:barChart',                     True),
        ('c:lineChart/c:varyColors',       True),
        ('c:pieChart/c:varyColors{val=0}', False),
    ])
    def vary_by_categories_get_fixture(self, request):
        xChart_cxml, expected_value = request.param
        plot = Plot(element(xChart_cxml), None)
        return plot, expected_value

    @pytest.fixture(params=[
        ('c:barChart', False, 'c:barChart/c:varyColors{val=0}'),
        ('c:barChart', True,  'c:barChart/c:varyColors'),
        ('c:lineChart/c:varyColors{val=0}', True,
         'c:lineChart/c:varyColors'),
        ('c:pieChart/c:varyColors{val=1}',  False,
         'c:pieChart/c:varyColors{val=0}'),
    ])
    def vary_by_categories_set_fixture(self, request):
        xChart_cxml, new_value, expected_xChart_cxml = request.param
        plot = Plot(element(xChart_cxml), None)
        expected_xml = xml(expected_xChart_cxml)
        return plot, new_value, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def chart_(self, request):
        return instance_mock(request, Chart)

    @pytest.fixture
    def DataLabels_(self, request, data_labels_):
        return class_mock(
            request, 'pptx.chart.plot.DataLabels',
            return_value=data_labels_
        )

    @pytest.fixture
    def data_labels_(self, request):
        return instance_mock(request, DataLabels)

    @pytest.fixture
    def SeriesCollection_(self, request, series_collection_):
        return class_mock(
            request, 'pptx.chart.plot.SeriesCollection',
            return_value=series_collection_
        )

    @pytest.fixture
    def series_collection_(self, request):
        return instance_mock(request, SeriesCollection)


class DescribeBarPlot(object):

    def it_knows_its_gap_width(self, gap_width_get_fixture):
        bar_plot, expected_value = gap_width_get_fixture
        assert bar_plot.gap_width == expected_value

    def it_can_change_its_gap_width(self, gap_width_set_fixture):
        bar_plot, new_value, expected_xml = gap_width_set_fixture
        bar_plot.gap_width = new_value
        assert bar_plot._element.xml == expected_xml

    def it_knows_how_much_it_overlaps_the_adjacent_bar(
            self, overlap_get_fixture):
        bar_plot, expected_value = overlap_get_fixture
        assert bar_plot.overlap == expected_value

    def it_can_change_how_much_it_overlaps_the_adjacent_bar(
            self, overlap_set_fixture):
        bar_plot, new_value, expected_xml = overlap_set_fixture
        bar_plot.overlap = new_value
        assert bar_plot._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:barChart',                      150),
        ('c:barChart/c:gapWidth',           150),
        ('c:barChart/c:gapWidth{val=175}',  175),
        ('c:barChart/c:gapWidth{val=042%}',  42),
    ])
    def gap_width_get_fixture(self, request):
        barChart_cxml, expected_value = request.param
        bar_plot = BarPlot(element(barChart_cxml), None)
        return bar_plot, expected_value

    @pytest.fixture(params=[
        ('c:barChart',  42, 'c:barChart/c:gapWidth{val=42}'),
        ('c:barChart', 150, 'c:barChart/c:gapWidth'),
        ('c:barChart/c:gapWidth{val=200%}', 300,
         'c:barChart/c:gapWidth{val=300}'),
    ])
    def gap_width_set_fixture(self, request):
        barChart_cxml, new_value, expected_barChart_cxml = request.param
        bar_plot = BarPlot(element(barChart_cxml), None)
        expected_xml = xml(expected_barChart_cxml)
        return bar_plot, new_value, expected_xml

    @pytest.fixture(params=[
        ('c:barChart',                        0),
        ('c:barChart/c:overlap',              0),
        ('c:barChart/c:overlap{val=42}',     42),
        ('c:barChart/c:overlap{val=-42}',   -42),
        ('c:barChart/c:overlap{val=-042%}', -42),
    ])
    def overlap_get_fixture(self, request):
        barChart_cxml, expected_value = request.param
        bar_plot = BarPlot(element(barChart_cxml), None)
        return bar_plot, expected_value

    @pytest.fixture(params=[
        ('c:barChart',                   42, 'c:barChart/c:overlap{val=42}'),
        ('c:barChart/c:overlap{val=42}', 24, 'c:barChart/c:overlap{val=24}'),
        ('c:barChart/c:overlap{val=42}', -5, 'c:barChart/c:overlap{val=-5}'),
        ('c:barChart/c:overlap{val=42}',  0, 'c:barChart'),
    ])
    def overlap_set_fixture(self, request):
        barChart_cxml, new_value, expected_barChart_cxml = request.param
        bar_plot = BarPlot(element(barChart_cxml), None)
        expected_xml = xml(expected_barChart_cxml)
        return bar_plot, new_value, expected_xml


class DescribeDataLabels(object):

    def it_provides_access_to_its_font(self, font_fixture):
        data_labels, Font_, defRPr, font_ = font_fixture
        font = data_labels.font
        Font_.assert_called_once_with(defRPr)
        assert font is font_

    def it_adds_a_txPr_to_help_font(self, txPr_fixture):
        data_labels, expected_xml = txPr_fixture
        data_labels.font
        assert data_labels._element.xml == expected_xml

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

    def it_knows_its_position(self, position_get_fixture):
        data_labels, expected_value = position_get_fixture
        assert data_labels.position == expected_value

    def it_can_change_its_position(self, position_set_fixture):
        data_labels, new_value, expected_xml = position_set_fixture
        data_labels.position = new_value
        assert data_labels._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def font_fixture(self, Font_, font_):
        dLbls = element('c:dLbls/c:txPr/a:p/a:pPr/a:defRPr')
        defRPr = dLbls.xpath('.//a:defRPr')[0]
        data_labels = DataLabels(dLbls)
        return data_labels, Font_, defRPr, font_

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

    @pytest.fixture(params=[
        ('c:dLbls',                       None),
        ('c:dLbls/c:dLblPos{val=inBase}', XL_LABEL_POSITION.INSIDE_BASE),
    ])
    def position_get_fixture(self, request):
        dLbls_cxml, expected_value = request.param
        data_labels = DataLabels(element(dLbls_cxml))
        return data_labels, expected_value

    @pytest.fixture(params=[
        ('c:dLbls',                       XL_LABEL_POSITION.INSIDE_BASE,
         'c:dLbls/c:dLblPos{val=inBase}'),
        ('c:dLbls/c:dLblPos{val=inBase}', XL_LABEL_POSITION.OUTSIDE_END,
         'c:dLbls/c:dLblPos{val=outEnd}'),
        ('c:dLbls/c:dLblPos{val=inBase}', None, 'c:dLbls'),
        ('c:dLbls',                       None, 'c:dLbls'),
    ])
    def position_set_fixture(self, request):
        dLbls_cxml, new_value, expected_dLbls_cxml = request.param
        data_labels = DataLabels(element(dLbls_cxml))
        expected_xml = xml(expected_dLbls_cxml)
        return data_labels, new_value, expected_xml

    @pytest.fixture(params=[
        ('c:dLbls{a:b=c}',
         'c:dLbls{a:b=c}/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)'),
        ('c:dLbls{a:b=c}/c:txPr/(a:bodyPr,a:p)',
         'c:dLbls{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr/a:defRPr)'),
        ('c:dLbls{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr)',
         'c:dLbls{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr/a:defRPr)'),
    ])
    def txPr_fixture(self, request):
        dLbls_cxml, expected_cxml = request.param
        data_labels = DataLabels(element(dLbls_cxml))
        expected_xml = xml(expected_cxml)
        return data_labels, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Font_(self, request, font_):
        return class_mock(
            request, 'pptx.chart.plot.Font', return_value=font_
        )

    @pytest.fixture
    def font_(self, request):
        return instance_mock(request, Font)


class DescribePlotFactory(object):

    def it_contructs_a_plot_object_from_a_plot_element(self, call_fixture):
        xChart, chart_, PlotClass_, plot_ = call_fixture
        plot = PlotFactory(xChart, chart_)
        PlotClass_.assert_called_once_with(xChart, chart_)
        assert plot is plot_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:areaChart',   AreaPlot),
        ('c:area3DChart', Area3DPlot),
        ('c:barChart',    BarPlot),
        ('c:lineChart',   LinePlot),
        ('c:pieChart',    PiePlot),
    ])
    def call_fixture(self, request, chart_):
        xChart_cxml, PlotCls = request.param
        plot_ = instance_mock(request, PlotCls, name='plot_')
        class_spec = 'pptx.chart.plot.%s' % PlotCls.__name__
        PlotClass_ = class_mock(request, class_spec, return_value=plot_)
        xChart = element(xChart_cxml)
        return xChart, chart_, PlotClass_, plot_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def chart_(self, request):
        return instance_mock(request, Chart)


class DescribePlotTypeInspector(object):

    def it_can_determine_the_chart_type_of_a_plot(self, chart_type_fixture):
        plot, expected_chart_type = chart_type_fixture
        chart_type = PlotTypeInspector.chart_type(plot)
        assert chart_type is expected_chart_type

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:areaChart',                            XL.AREA),
        ('c:areaChart/c:grouping',                 XL.AREA),
        ('c:areaChart/c:grouping{val=standard}',   XL.AREA),
        ('c:areaChart/c:grouping{val=stacked}',    XL.AREA_STACKED),
        ('c:areaChart/c:grouping{val=percentStacked}',
         XL.AREA_STACKED_100),

        ('c:area3DChart',                          XL.THREE_D_AREA),
        ('c:area3DChart/c:grouping',               XL.THREE_D_AREA),
        ('c:area3DChart/c:grouping{val=standard}', XL.THREE_D_AREA),
        ('c:area3DChart/c:grouping{val=stacked}',  XL.THREE_D_AREA_STACKED),
        ('c:area3DChart/c:grouping{val=percentStacked}',
         XL.THREE_D_AREA_STACKED_100),

        ('c:barChart/c:barDir',                       XL.COLUMN_CLUSTERED),
        ('c:barChart/c:barDir{val=col}',              XL.COLUMN_CLUSTERED),
        ('c:barChart/(c:barDir{val=col},c:grouping)', XL.COLUMN_CLUSTERED),
        ('c:barChart/(c:barDir{val=col},c:grouping{val=clustered})',
         XL.COLUMN_CLUSTERED),
        ('c:barChart/(c:barDir{val=col},c:grouping{val=stacked})',
         XL.COLUMN_STACKED),
        ('c:barChart/(c:barDir{val=col},c:grouping{val=percentStacked})',
         XL.COLUMN_STACKED_100),

        ('c:barChart/c:barDir{val=bar}',              XL.BAR_CLUSTERED),
        ('c:barChart/(c:barDir{val=bar},c:grouping)', XL.BAR_CLUSTERED),
        ('c:barChart/(c:barDir{val=bar},c:grouping{val=clustered})',
         XL.BAR_CLUSTERED),
        ('c:barChart/(c:barDir{val=bar},c:grouping{val=stacked})',
         XL.BAR_STACKED),
        ('c:barChart/(c:barDir{val=bar},c:grouping{val=percentStacked})',
         XL.BAR_STACKED_100),

        ('c:lineChart',                          XL.LINE_MARKERS),
        ('c:lineChart/c:grouping',               XL.LINE_MARKERS),
        ('c:lineChart/c:grouping{val=standard}', XL.LINE_MARKERS),
        ('c:lineChart/c:grouping{val=stacked}',  XL.LINE_MARKERS_STACKED),
        ('c:lineChart/c:grouping{val=percentStacked}',
         XL.LINE_MARKERS_STACKED_100),
        ('c:lineChart/c:ser/c:marker/c:symbol{val=none}', XL.LINE),
        ('c:lineChart/(c:grouping{val=stacked},c:ser/c:marker/c:symbol{val=n'
         'one})', XL.LINE_STACKED),
        ('c:lineChart/(c:grouping{val=percentStacked},c:ser/c:marker/c:symbo'
         'l{val=none})', XL.LINE_STACKED_100),

        ('c:pieChart',                           XL.PIE),
        ('c:pieChart/c:ser/c:explosion{val=25}', XL.PIE_EXPLODED),
    ])
    def chart_type_fixture(self, request):
        xChart_cxml, expected_chart_type = request.param
        plot = PlotFactory(element(xChart_cxml), None)
        return plot, expected_chart_type
