# encoding: utf-8

"""
Test suite for pptx.chart.xmlwriter module
"""

from __future__ import absolute_import, print_function, unicode_literals

from itertools import islice

import pytest

from pptx.chart.data import (
    _BaseChartData, _BaseSeriesData, BubbleChartData, CategoryChartData,
    XyChartData
)
from pptx.chart.xmlwriter import (
    _BarChartXmlWriter, _BaseSeriesXmlRewriter, _BubbleChartXmlWriter,
    _BubbleSeriesXmlRewriter, _CategorySeriesXmlRewriter, ChartXmlWriter,
    _LineChartXmlWriter, _PieChartXmlWriter, _PieSeriesXmlRewriter,
    _RadarChartXmlWriter, SeriesXmlRewriterFactory, _XyChartXmlWriter,
    _XySeriesXmlRewriter
)
from pptx.enum.chart import XL_CHART_TYPE

from ..unitutil import count
from ..unitutil.cxml import element
from ..unitutil.file import snippet_text
from ..unitutil.mock import call, class_mock, instance_mock, method_mock


class DescribeChartXmlWriter(object):

    def it_contructs_an_xml_writer_for_a_chart_type(self, call_fixture):
        chart_type, series_seq_, XmlWriterClass_, xml_writer_ = call_fixture
        xml_writer = ChartXmlWriter(chart_type, series_seq_)
        XmlWriterClass_.assert_called_once_with(chart_type, series_seq_)
        assert xml_writer is xml_writer_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('BAR_CLUSTERED',                _BarChartXmlWriter),
        ('BAR_STACKED_100',              _BarChartXmlWriter),
        ('BUBBLE',                       _BubbleChartXmlWriter),
        ('BUBBLE_THREE_D_EFFECT',        _BubbleChartXmlWriter),
        ('COLUMN_CLUSTERED',             _BarChartXmlWriter),
        ('LINE',                         _LineChartXmlWriter),
        ('PIE',                          _PieChartXmlWriter),
        ('RADAR',                        _RadarChartXmlWriter),
        ('RADAR_FILLED',                 _RadarChartXmlWriter),
        ('RADAR_MARKERS',                _RadarChartXmlWriter),
        ('XY_SCATTER',                   _XyChartXmlWriter),
        ('XY_SCATTER_LINES',             _XyChartXmlWriter),
        ('XY_SCATTER_LINES_NO_MARKERS',  _XyChartXmlWriter),
        ('XY_SCATTER_SMOOTH',            _XyChartXmlWriter),
        ('XY_SCATTER_SMOOTH_NO_MARKERS', _XyChartXmlWriter),
    ])
    def call_fixture(self, request, series_seq_):
        chart_type_member, XmlWriterClass = request.param
        xml_writer_ = instance_mock(request, XmlWriterClass)
        class_spec = 'pptx.chart.xmlwriter.%s' % XmlWriterClass.__name__
        XmlWriterClass_ = class_mock(
            request, class_spec, return_value=xml_writer_
        )
        chart_type = getattr(XL_CHART_TYPE, chart_type_member)
        return chart_type, series_seq_, XmlWriterClass_, xml_writer_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def series_seq_(self, request):
        return instance_mock(request, tuple)


class DescribeSeriesXmlRewriterFactory(object):

    def it_contructs_an_xml_rewriter_for_a_chart_type(self, call_fixture):
        chart_type, chart_data_, XmlRewriterClass_, xml_rewriter_ = (
            call_fixture
        )

        xml_rewriter = SeriesXmlRewriterFactory(chart_type, chart_data_)

        XmlRewriterClass_.assert_called_once_with(chart_data_)
        assert xml_rewriter is xml_rewriter_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('BAR_CLUSTERED',                _CategorySeriesXmlRewriter),
        ('BAR_OF_PIE',                   _PieSeriesXmlRewriter),
        ('BUBBLE',                       _BubbleSeriesXmlRewriter),
        ('BUBBLE_THREE_D_EFFECT',        _BubbleSeriesXmlRewriter),
        ('DOUGHNUT',                     _PieSeriesXmlRewriter),
        ('DOUGHNUT_EXPLODED',            _PieSeriesXmlRewriter),
        ('PIE',                          _PieSeriesXmlRewriter),
        ('PIE_EXPLODED',                 _PieSeriesXmlRewriter),
        ('PIE_OF_PIE',                   _PieSeriesXmlRewriter),
        ('THREE_D_PIE',                  _PieSeriesXmlRewriter),
        ('THREE_D_PIE_EXPLODED',         _PieSeriesXmlRewriter),
        ('XY_SCATTER',                   _XySeriesXmlRewriter),
        ('XY_SCATTER_LINES',             _XySeriesXmlRewriter),
        ('XY_SCATTER_LINES_NO_MARKERS',  _XySeriesXmlRewriter),
        ('XY_SCATTER_SMOOTH',            _XySeriesXmlRewriter),
        ('XY_SCATTER_SMOOTH_NO_MARKERS', _XySeriesXmlRewriter),
    ])
    def call_fixture(self, request, chart_data_):
        chart_type_member, rewriter_cls = request.param
        chart_type = getattr(XL_CHART_TYPE, chart_type_member)
        xml_rewriter_ = instance_mock(request, rewriter_cls)
        XmlRewriterClass_ = class_mock(
            request, 'pptx.chart.xmlwriter.%s' % rewriter_cls.__name__,
            return_value=xml_rewriter_
        )
        return chart_type, chart_data_, XmlRewriterClass_, xml_rewriter_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def chart_data_(self, request):
        return instance_mock(request, _BaseChartData)


class Describe_BarChartXmlWriter(object):

    def it_can_generate_xml_for_bar_type_charts(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        assert xml_writer.xml == expected_xml

    def it_can_generate_xml_for_multi_level_cat_charts(self, multi_fixture):
        xml_writer, expected_xml = multi_fixture
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def multi_fixture(self):
        chart_data = CategoryChartData()

        WEST = chart_data.add_category('WEST')
        WEST.add_sub_category('SF')
        WEST.add_sub_category('LA')
        EAST = chart_data.add_category('EAST')
        EAST.add_sub_category('NY')
        EAST.add_sub_category('NJ')

        chart_data.add_series('Series 1', (1, 2, None, 4))
        chart_data.add_series('Series 2', (5, None, 7, 8))

        xml_writer = _BarChartXmlWriter(
            XL_CHART_TYPE.BAR_CLUSTERED, chart_data
        )
        expected_xml = snippet_text('4x2-multi-cat-bar')
        return xml_writer, expected_xml

    @pytest.fixture(params=[
        ('BAR_CLUSTERED',    2, 2, '2x2-bar-clustered'),
        ('BAR_STACKED_100',  2, 2, '2x2-bar-stacked-100'),
        ('COLUMN_CLUSTERED', 2, 2, '2x2-column-clustered'),
    ])
    def xml_fixture(self, request):
        enum_member, cat_count, ser_count, snippet_name = request.param
        chart_type = getattr(XL_CHART_TYPE, enum_member)
        chart_data = make_category_chart_data(cat_count, ser_count)
        xml_writer = _BarChartXmlWriter(chart_type, chart_data)
        expected_xml = snippet_text(snippet_name)
        return xml_writer, expected_xml


class Describe_BubbleChartXmlWriter(object):

    def it_can_generate_xml_for_bubble_charts(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('BUBBLE',                2, 3, '2x3-bubble'),
        ('BUBBLE_THREE_D_EFFECT', 2, 3, '2x3-bubble-3d'),
    ])
    def xml_fixture(self, request):
        enum_member, ser_count, point_count, snippet_name = request.param
        chart_type = getattr(XL_CHART_TYPE, enum_member)
        chart_data = make_bubble_chart_data(ser_count, point_count)
        xml_writer = _BubbleChartXmlWriter(chart_type, chart_data)
        expected_xml = snippet_text(snippet_name)
        return xml_writer, expected_xml


class Describe_LineChartXmlWriter(object):

    def it_can_generate_xml_for_a_line_chart(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def xml_fixture(self, request):
        series_data_seq = make_category_chart_data(cat_count=2, ser_count=2)
        xml_writer = _LineChartXmlWriter(
            XL_CHART_TYPE.LINE, series_data_seq
        )
        expected_xml = snippet_text('2x2-line')
        return xml_writer, expected_xml


class Describe_PieChartXmlWriter(object):

    def it_can_generate_xml_for_a_pie_chart(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def xml_fixture(self, request):
        series_data_seq = make_category_chart_data(cat_count=3, ser_count=1)
        xml_writer = _PieChartXmlWriter(XL_CHART_TYPE.PIE, series_data_seq)
        expected_xml = snippet_text('3x1-pie')
        return xml_writer, expected_xml


class Describe_RadarChartXmlWriter(object):

    def it_can_generate_xml_for_a_radar_chart(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        print(xml_writer.xml)
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def xml_fixture(self, request):
        series_data_seq = make_category_chart_data(cat_count=5, ser_count=2)
        xml_writer = _RadarChartXmlWriter(
            XL_CHART_TYPE.RADAR, series_data_seq
        )
        expected_xml = snippet_text('2x5-radar')
        return xml_writer, expected_xml


class Describe_XyChartXmlWriter(object):

    def it_can_generate_xml_for_xy_charts(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('XY_SCATTER',                   2, 3, '2x3-xy'),
        ('XY_SCATTER_LINES',             2, 3, '2x3-xy-lines'),
        ('XY_SCATTER_LINES_NO_MARKERS',  2, 3, '2x3-xy-lines-no-markers'),
        ('XY_SCATTER_SMOOTH',            2, 3, '2x3-xy-smooth'),
        ('XY_SCATTER_SMOOTH_NO_MARKERS', 2, 3, '2x3-xy-smooth-no-markers'),
    ])
    def xml_fixture(self, request):
        enum_member, ser_count, point_count, snippet_name = request.param
        chart_type = getattr(XL_CHART_TYPE, enum_member)
        chart_data = make_xy_chart_data(ser_count, point_count)
        xml_writer = _XyChartXmlWriter(chart_type, chart_data)
        expected_xml = snippet_text(snippet_name)
        return xml_writer, expected_xml


class Describe_BaseSeriesXmlRewriter(object):

    def it_can_replace_series_data(self, replace_fixture):
        rewriter, chartSpace, ser_count, calls = replace_fixture
        rewriter.replace_series_data(chartSpace)
        rewriter._adjust_ser_count.assert_called_once_with(
            rewriter, chartSpace, ser_count
        )
        assert rewriter._rewrite_ser_data.call_args_list == calls

    def it_adjusts_the_ser_count_to_help(self, adjust_fixture):
        rewriter, chartSpace, ser_count, add_calls, trim_calls, sers = (
            adjust_fixture
        )
        _sers = rewriter._adjust_ser_count(chartSpace, ser_count)
        assert rewriter._add_cloned_sers.call_args_list == add_calls
        assert rewriter._trim_ser_count_by.call_args_list == trim_calls
        assert _sers == sers

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (3, True, False),
        (2, False, False),
        (1, False, True),
    ])
    def adjust_fixture(self, request, _add_cloned_sers_, _trim_ser_count_by_):
        ser_count, add, trim = request.param
        rewriter = _BaseSeriesXmlRewriter(None)
        chartSpace = element(
            'c:chartSpace/(c:ser/(c:idx{val=3},c:order{val=1}),c:ser/(c:idx{'
            'val=1},c:order{val=3}))'
        )
        sers = chartSpace.sers
        add_calls = [call(rewriter, chartSpace, 1)] if add else []
        trim_calls = [call(rewriter, chartSpace, 1)] if trim else []
        return rewriter, chartSpace, ser_count, add_calls, trim_calls, sers

    @pytest.fixture
    def replace_fixture(
            self, request, chart_data_, _adjust_ser_count_,
            _rewrite_ser_data_):
        rewriter = _BaseSeriesXmlRewriter(chart_data_)
        chartSpace = element('c:chartSpace/(c:ser,c:ser)')
        sers = chartSpace.xpath('c:ser')
        ser_count = len(sers)
        series_datas = [
            instance_mock(request, _BaseSeriesData),
            instance_mock(request, _BaseSeriesData),
        ]
        calls = [
            call(rewriter, sers[0], series_datas[0]),
            call(rewriter, sers[1], series_datas[1])
        ]
        chart_data_.__len__.return_value = len(series_datas)
        chart_data_.__iter__.return_value = iter(series_datas)
        _adjust_ser_count_.return_value = sers
        return rewriter, chartSpace, ser_count, calls

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _add_cloned_sers_(self, request):
        return method_mock(
            request, _BaseSeriesXmlRewriter, '_add_cloned_sers',
            autospec=True
        )

    @pytest.fixture
    def _adjust_ser_count_(self, request):
        return method_mock(
            request, _BaseSeriesXmlRewriter, '_adjust_ser_count',
            autospec=True
        )

    @pytest.fixture
    def chart_data_(self, request):
        return instance_mock(request, _BaseChartData)

    @pytest.fixture
    def _rewrite_ser_data_(self, request):
        return method_mock(
            request, _BaseSeriesXmlRewriter, '_rewrite_ser_data',
            autospec=True
        )

    @pytest.fixture
    def _trim_ser_count_by_(self, request):
        return method_mock(
            request, _BaseSeriesXmlRewriter, '_trim_ser_count_by',
            autospec=True
        )


# helpers ------------------------------------------------------------

def make_bubble_chart_data(ser_count, point_count):
    """
    Return an |BubbleChartData| object populated with *ser_count* series,
    each having *point_count* data points.
    """
    points = (
        (1.1, 11.1, 10.0), (2.1, 12.1, 20.0), (3.1, 13.1, 30.0),
        (1.2, 11.2, 40.0), (2.2, 12.2, 50.0), (3.2, 13.2, 60.0),
    )
    chart_data = BubbleChartData()
    for i in range(ser_count):
        series_label = 'Series %d' % (i+1)
        series = chart_data.add_series(series_label)
        for j in range(point_count):
            point_idx = (i * point_count) + j
            x, y, size = points[point_idx]
            series.add_data_point(x, y, size)
    return chart_data


def make_category_chart_data(cat_count, ser_count):
    """
    Return a |CategoryChartData| instance populated with *cat_count*
    categories and *ser_count* series. Values are auto-generated.
    """
    category_names = ('Foo', 'Bar', 'Baz', 'Boo', 'Far', 'Faz')
    point_values = count(1.1, 1.1)
    chart_data = CategoryChartData()
    chart_data.categories = category_names[:cat_count]
    for idx in range(ser_count):
        series_title = 'Series %d' % (idx+1)
        series_values = tuple(islice(point_values, cat_count))
        series_values = [round(x*10)/10.0 for x in series_values]
        chart_data.add_series(series_title, series_values)
    return chart_data


def make_xy_chart_data(ser_count, point_count):
    """
    Return an |XyChartData| object populated with *ser_count* series each
    having *point_count* data points. Values are auto-generated.
    """
    points = (
        (1.1, 11.1), (2.1, 12.1), (3.1, 13.1),
        (1.2, 11.2), (2.2, 12.2), (3.2, 13.2),
    )
    chart_data = XyChartData()
    for i in range(ser_count):
        series_label = 'Series %d' % (i+1)
        series = chart_data.add_series(series_label)
        for j in range(point_count):
            point_idx = (i * point_count) + j
            x, y = points[point_idx]
            series.add_data_point(x, y)
    return chart_data
