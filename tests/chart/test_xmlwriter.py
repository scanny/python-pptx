# encoding: utf-8

"""
Test suite for pptx.chart.xmlwriter module
"""

from __future__ import absolute_import, print_function, unicode_literals

from datetime import date
from itertools import islice

import pytest

from pptx.chart.data import (
    _BaseChartData, _BaseSeriesData, BubbleChartData, CategoryChartData,
    CategorySeriesData, XyChartData
)
from pptx.chart.xmlwriter import (
    _AreaChartXmlWriter, _BarChartXmlWriter, _BaseSeriesXmlRewriter,
    _BubbleChartXmlWriter, _BubbleSeriesXmlRewriter, _BubbleSeriesXmlWriter,
    _CategorySeriesXmlRewriter, _CategorySeriesXmlWriter, ChartXmlWriter,
    _DoughnutChartXmlWriter, _LineChartXmlWriter, _PieChartXmlWriter,
    _RadarChartXmlWriter, SeriesXmlRewriterFactory, _XyChartXmlWriter,
    _XySeriesXmlRewriter, _XySeriesXmlWriter
)
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml import parse_xml

from ..unitutil import count
from ..unitutil.cxml import element, xml
from ..unitutil.file import snippet_seq, snippet_text
from ..unitutil.mock import call, class_mock, instance_mock, method_mock


class DescribeChartXmlWriter(object):

    def it_contructs_an_xml_writer_for_a_chart_type(self, call_fixture):
        chart_type, series_seq_, XmlWriterClass_, xml_writer_ = call_fixture
        xml_writer = ChartXmlWriter(chart_type, series_seq_)
        XmlWriterClass_.assert_called_once_with(chart_type, series_seq_)
        assert xml_writer is xml_writer_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('AREA',                         _AreaChartXmlWriter),
        ('AREA_STACKED',                 _AreaChartXmlWriter),
        ('AREA_STACKED_100',             _AreaChartXmlWriter),
        ('BAR_CLUSTERED',                _BarChartXmlWriter),
        ('BAR_STACKED',                  _BarChartXmlWriter),
        ('BAR_STACKED_100',              _BarChartXmlWriter),
        ('BUBBLE',                       _BubbleChartXmlWriter),
        ('BUBBLE_THREE_D_EFFECT',        _BubbleChartXmlWriter),
        ('COLUMN_CLUSTERED',             _BarChartXmlWriter),
        ('COLUMN_STACKED',               _BarChartXmlWriter),
        ('COLUMN_STACKED_100',           _BarChartXmlWriter),
        ('DOUGHNUT',                     _DoughnutChartXmlWriter),
        ('DOUGHNUT_EXPLODED',            _DoughnutChartXmlWriter),
        ('LINE',                         _LineChartXmlWriter),
        ('LINE_MARKERS',                 _LineChartXmlWriter),
        ('LINE_MARKERS_STACKED',         _LineChartXmlWriter),
        ('LINE_MARKERS_STACKED_100',     _LineChartXmlWriter),
        ('LINE_STACKED',                 _LineChartXmlWriter),
        ('LINE_STACKED_100',             _LineChartXmlWriter),
        ('PIE',                          _PieChartXmlWriter),
        ('PIE_EXPLODED',                 _PieChartXmlWriter),
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
        ('BUBBLE',                       _BubbleSeriesXmlRewriter),
        ('BUBBLE_THREE_D_EFFECT',        _BubbleSeriesXmlRewriter),
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


class Describe_AreaChartXmlWriter(object):

    def it_can_generate_xml_for_area_type_charts(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('AREA',             2, 2, str,   '2x2-area'),
        ('AREA',             2, 2, date,  '2x2-area-date'),
        ('AREA',             2, 2, float, '2x2-area-float'),
        ('AREA_STACKED',     2, 2, str,   '2x2-area-stacked'),
        ('AREA_STACKED_100', 2, 2, str,   '2x2-area-stacked-100'),
    ])
    def xml_fixture(self, request):
        member, cat_count, ser_count, cat_type, snippet_name = request.param
        chart_type = getattr(XL_CHART_TYPE, member)
        chart_data = make_category_chart_data(cat_count, cat_type, ser_count)
        xml_writer = _AreaChartXmlWriter(chart_type, chart_data)
        expected_xml = snippet_text(snippet_name)
        return xml_writer, expected_xml


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
        ('BAR_CLUSTERED',      2, 2, str,   '2x2-bar-clustered'),
        ('BAR_CLUSTERED',      2, 2, date,  '2x2-bar-clustered-date'),
        ('BAR_CLUSTERED',      2, 2, float, '2x2-bar-clustered-float'),
        ('BAR_STACKED',        2, 2, str,   '2x2-bar-stacked'),
        ('BAR_STACKED_100',    2, 2, str,   '2x2-bar-stacked-100'),
        ('COLUMN_CLUSTERED',   2, 2, str,   '2x2-column-clustered'),
        ('COLUMN_STACKED',     2, 2, str,   '2x2-column-stacked'),
        ('COLUMN_STACKED_100', 2, 2, str,   '2x2-column-stacked-100'),
    ])
    def xml_fixture(self, request):
        member, cat_count, ser_count, cat_type, snippet_name = request.param
        chart_type = getattr(XL_CHART_TYPE, member)
        chart_data = make_category_chart_data(cat_count, cat_type, ser_count)
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


class Describe_DoughnutChartXmlWriter(object):

    def it_can_generate_xml_for_doughnut_type_charts(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('DOUGHNUT',          3, 2, '3x2-doughnut'),
        ('DOUGHNUT_EXPLODED', 3, 2, '3x2-doughnut-exploded'),
    ])
    def xml_fixture(self, request):
        enum_member, cat_count, ser_count, snippet_name = request.param
        chart_type = getattr(XL_CHART_TYPE, enum_member)
        chart_data = make_category_chart_data(cat_count, str, ser_count)
        xml_writer = _DoughnutChartXmlWriter(chart_type, chart_data)
        expected_xml = snippet_text(snippet_name)
        return xml_writer, expected_xml


class Describe_LineChartXmlWriter(object):

    def it_can_generate_xml_for_a_line_chart(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('LINE',                     2, 2, str,   '2x2-line'),
        ('LINE',                     2, 2, date,  '2x2-line-date'),
        ('LINE',                     2, 2, float, '2x2-line-float'),
        ('LINE_MARKERS',             2, 2, str,   '2x2-line-markers'),
        ('LINE_MARKERS_STACKED',     2, 2, str,   '2x2-line-markers-stacked'),
        ('LINE_MARKERS_STACKED_100', 2, 2, str,
         '2x2-line-markers-stacked-100'),
        ('LINE_STACKED',             2, 2, str,   '2x2-line-stacked'),
        ('LINE_STACKED_100',         2, 2, str,   '2x2-line-stacked-100'),
    ])
    def xml_fixture(self, request):
        member, cat_count, ser_count, cat_type, snippet_name = request.param
        chart_type = getattr(XL_CHART_TYPE, member)
        chart_data = make_category_chart_data(cat_count, cat_type, ser_count)
        xml_writer = _LineChartXmlWriter(chart_type, chart_data)
        expected_xml = snippet_text(snippet_name)
        return xml_writer, expected_xml


class Describe_PieChartXmlWriter(object):

    def it_can_generate_xml_for_a_pie_chart(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('PIE',          3, 1, '3x1-pie'),
        ('PIE_EXPLODED', 3, 1, '3x1-pie-exploded'),
    ])
    def xml_fixture(self, request):
        enum_member, cat_count, ser_count, snippet_name = request.param
        chart_type = getattr(XL_CHART_TYPE, enum_member)
        chart_data = make_category_chart_data(cat_count, str, ser_count)
        xml_writer = _PieChartXmlWriter(chart_type, chart_data)
        expected_xml = snippet_text(snippet_name)
        return xml_writer, expected_xml


class Describe_RadarChartXmlWriter(object):

    def it_can_generate_xml_for_a_radar_chart(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        print(xml_writer.xml)
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def xml_fixture(self, request):
        series_data_seq = make_category_chart_data(
            cat_count=5, cat_type=str, ser_count=2
        )
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


class Describe_BubbleSeriesXmlWriter(object):

    def it_knows_its_bubbleSize_XML(self, bubbleSize_fixture):
        xml_writer, expected_xml = bubbleSize_fixture
        bubbleSize = xml_writer.bubbleSize
        assert bubbleSize.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (None, None, 'General'),
        (None, 42,   '42'),
        (24,   None, '24'),
        (24,   42,   '42'),
    ])
    def bubbleSize_fixture(self, request):
        cht_numfmt, ser_numfmt, expected_numfmt = request.param
        expected_xml = xml(
            'c:bubbleSize/c:numRef/(c:f"Sheet1!$C$2:$C$1",c:numCache/(c:form'
            'atCode"%s",c:ptCount{val=0}))' % expected_numfmt
        )
        chart_number_format = () if cht_numfmt is None else (cht_numfmt,)
        series_number_format = () if ser_numfmt is None else (ser_numfmt,)
        chart_data = BubbleChartData(*chart_number_format)
        series_data = chart_data.add_series(None, *series_number_format)
        xml_writer = _BubbleSeriesXmlWriter(series_data)
        return xml_writer, expected_xml


class Describe_CategorySeriesXmlWriter(object):

    def it_knows_its_val_XML(self, val_fixture):
        xml_writer, expected_xml = val_fixture
        val = xml_writer.val
        assert val.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def val_fixture(self, series_data_):
        values_ref, number_format = 'Sheet1!$B$2:$B$1', 'Foobar'
        expected_xml = xml(
            'c:val/c:numRef/(c:f"%s",c:numCache/(c:formatCode"%s",c:ptCount{'
            'val=0}))' % (values_ref, number_format)
        )
        xml_writer = _CategorySeriesXmlWriter(series_data_)
        series_data_.values_ref = values_ref
        series_data_.number_format = number_format
        return xml_writer, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def series_data_(self, request):
        return instance_mock(request, CategorySeriesData)


class Describe_XySeriesXmlWriter(object):

    def it_knows_its_xVal_XML(self, xVal_fixture):
        xml_writer, expected_xml = xVal_fixture
        xVal = xml_writer.xVal
        assert xVal.xml == expected_xml

    def it_knows_its_yVal_XML(self, yVal_fixture):
        xml_writer, expected_xml = yVal_fixture
        yVal = xml_writer.yVal
        assert yVal.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (None, None, 'General'),
        (None, 42,   '42'),
        (24,   None, '24'),
        (24,   42,   '42'),
    ])
    def xVal_fixture(self, request):
        cht_numfmt, ser_numfmt, expected_numfmt = request.param
        expected_xml = xml(
            'c:xVal/c:numRef/(c:f"Sheet1!$A$2:$A$1",c:numCache/(c:formatCode'
            '"%s",c:ptCount{val=0}))' % expected_numfmt
        )
        chart_number_format = () if cht_numfmt is None else (cht_numfmt,)
        series_number_format = () if ser_numfmt is None else (ser_numfmt,)
        chart_data = XyChartData(*chart_number_format)
        series_data = chart_data.add_series(None, *series_number_format)
        xml_writer = _XySeriesXmlWriter(series_data)
        return xml_writer, expected_xml

    @pytest.fixture(params=[
        (None, None, 'General'),
        (None, 42,   '42'),
        (24,   None, '24'),
        (24,   42,   '42'),
    ])
    def yVal_fixture(self, request):
        cht_numfmt, ser_numfmt, expected_numfmt = request.param
        expected_xml = xml(
            'c:yVal/c:numRef/(c:f"Sheet1!$B$2:$B$1",c:numCache/(c:formatCode'
            '"%s",c:ptCount{val=0}))' % expected_numfmt
        )
        chart_number_format = () if cht_numfmt is None else (cht_numfmt,)
        series_number_format = () if ser_numfmt is None else (ser_numfmt,)
        chart_data = XyChartData(*chart_number_format)
        series_data = chart_data.add_series(None, *series_number_format)
        xml_writer = _XySeriesXmlWriter(series_data)
        return xml_writer, expected_xml


class Describe_BaseSeriesXmlRewriter(object):

    def it_can_replace_series_data(self, replace_fixture):
        rewriter, chartSpace, plotArea, ser_count, calls = replace_fixture
        rewriter.replace_series_data(chartSpace)
        rewriter._adjust_ser_count.assert_called_once_with(
            rewriter, plotArea, ser_count
        )
        assert rewriter._rewrite_ser_data.call_args_list == calls

    def it_adjusts_the_ser_count_to_help(self, adjust_fixture):
        rewriter, plotArea, new_ser_count, add_calls, trim_calls, sers = (
            adjust_fixture
        )
        rewriter._adjust_ser_count(plotArea, new_ser_count)
        assert rewriter._add_cloned_sers.call_args_list == add_calls
        assert rewriter._trim_ser_count_by.call_args_list == trim_calls

    def it_adds_cloned_sers_to_help(self, clone_fixture):
        rewriter, plotArea, count, expected_xml = clone_fixture
        rewriter._add_cloned_sers(plotArea, count)
        assert plotArea.xml == expected_xml

    def it_trims_sers_to_help(self, trim_fixture):
        rewriter, plotArea, count, expected_xml = trim_fixture
        rewriter._trim_ser_count_by(plotArea, count)
        assert plotArea.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (3, True, False),
        (2, False, False),
        (1, False, True),
    ])
    def adjust_fixture(self, request, _add_cloned_sers_, _trim_ser_count_by_):
        new_ser_count, add, trim = request.param
        rewriter = _BaseSeriesXmlRewriter(None)
        plotArea = element(
            'c:plotArea/c:barChart/(c:ser/(c:idx{val=3},c:order{val=1}),c:se'
            'r/(c:idx{val=1},c:order{val=3}))'
        )
        sers = plotArea.sers
        add_calls = [call(rewriter, plotArea, 1)] if add else []
        trim_calls = [call(rewriter, plotArea, 1)] if trim else []
        return rewriter, plotArea, new_ser_count, add_calls, trim_calls, sers

    @pytest.fixture(params=[
        ('c:plotArea/c:barChart/c:ser/(c:idx{val=3},c:order{val=1})', 1,
         'c:plotArea/c:barChart/(c:ser/(c:idx{val=3},c:order{val=1}),c:ser/('
         'c:idx{val=4},c:order{val=2}))'),
        ('c:plotArea/(c:barChart/c:ser/(c:idx{val=3},c:order{val=1}),c:lineC'
         'hart/(c:ser/(c:idx{val=4},c:order{val=7}),c:ser/(c:idx{val=1},c:or'
         'der{val=2})))', 1,
         'c:plotArea/(c:barChart/c:ser/(c:idx{val=3},c:order{val=1}),c:lineC'
         'hart/(c:ser/(c:idx{val=4},c:order{val=7}),c:ser/(c:idx{val=5},c:or'
         'der{val=8}),c:ser/(c:idx{val=1},c:order{val=2})))'),
    ])
    def clone_fixture(self, request):
        plotArea_cxml, count, expected_cxml = request.param
        rewriter = _BaseSeriesXmlRewriter(None)
        plotArea = element(plotArea_cxml)
        expected_xml = xml(expected_cxml)
        return rewriter, plotArea, count, expected_xml

    @pytest.fixture
    def replace_fixture(self, request, chart_data_, _adjust_ser_count_,
                        _rewrite_ser_data_):
        rewriter = _BaseSeriesXmlRewriter(chart_data_)
        chartSpace = element(
            'c:chartSpace/c:chart/c:plotArea/c:barChart/(c:ser/c:order{val=0'
            '},c:ser/c:order{val=1})'
        )
        plotArea = chartSpace.xpath('c:chart/c:plotArea')[0]
        sers = chartSpace.xpath('.//c:ser')
        ser_count = len(sers)
        series_datas = [
            instance_mock(request, _BaseSeriesData),
            instance_mock(request, _BaseSeriesData),
        ]
        calls = [
            call(rewriter, sers[0], series_datas[0], False),
            call(rewriter, sers[1], series_datas[1], False)
        ]
        chart_data_.__len__.return_value = len(series_datas)
        chart_data_.__iter__.return_value = iter(series_datas)
        return rewriter, chartSpace, plotArea, ser_count, calls

    @pytest.fixture(params=[
        # --single xChart, remove single ser--
        ('c:plotArea/c:barChart/(c:ser/(c:idx{val=3},c:order{val=4}),c:ser/('
         'c:idx{val=1},c:order{val=2}))', 1,
         'c:plotArea/c:barChart/(c:ser/(c:idx{val=1},c:order{val=2}))'),
        # --two xCharts, remove two sers (and emptied second xChart)--
        ('c:plotArea/(c:barChart/c:ser/(c:idx{val=3},c:order{val=1}),c:lineC'
         'hart/(c:ser/(c:idx{val=4},c:order{val=7}),c:ser/(c:idx{val=1},c:or'
         'der{val=2})))', 2,
         'c:plotArea/(c:barChart/c:ser/(c:idx{val=3},c:order{val=1}))'),
    ])
    def trim_fixture(self, request):
        plotArea_cxml, count, expected_cxml = request.param
        rewriter = _BaseSeriesXmlRewriter(None)
        plotArea = element(plotArea_cxml)
        expected_xml = xml(expected_cxml)
        return rewriter, plotArea, count, expected_xml

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


class Describe_BubbleSeriesXmlRewriter(object):

    def it_can_rewrite_a_ser_element(self, rewrite_fixture):
        rewriter, ser, series_data, expected_xml = rewrite_fixture
        rewriter._rewrite_ser_data(ser, series_data, None)
        assert ser.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def rewrite_fixture(self):
        ser_xml, expected_xml = snippet_seq('rewrite-ser')[4:6]

        chart_data = BubbleChartData()
        series_data = chart_data.add_series('Series 1')
        series_data.add_data_point(1, 2, 10)
        series_data.add_data_point(3, 4, 20)

        rewriter = _BubbleSeriesXmlRewriter(chart_data)
        ser = parse_xml(ser_xml)
        return rewriter, ser, series_data, expected_xml


class Describe_CategorySeriesXmlRewriter(object):

    def it_can_rewrite_a_ser_element(self, rewrite_fixture):
        rewriter, ser, series_data, expected_xml = rewrite_fixture
        rewriter._rewrite_ser_data(ser, series_data, False)
        assert ser.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (0, ('foo', 'bar')),
        (6, (42, 24)),
        (8, (date(2016, 12, 22), date(2016, 12, 23))),
    ])
    def rewrite_fixture(self, request):
        snippet_offset, categories = request.param
        chart_data = CategoryChartData()
        chart_data.categories = categories
        rewriter = _CategorySeriesXmlRewriter(chart_data)
        snippets = snippet_seq('rewrite-ser')
        ser_xml = snippets[snippet_offset]
        ser = parse_xml(ser_xml)
        series_data = chart_data.add_series('Series 1', (1, 2))
        expected_xml = snippets[snippet_offset+1]
        return rewriter, ser, series_data, expected_xml


class Describe_XySeriesXmlRewriter(object):

    def it_can_rewrite_a_ser_element(self, rewrite_fixture):
        rewriter, ser, series_data, expected_xml = rewrite_fixture
        rewriter._rewrite_ser_data(ser, series_data, None)
        assert ser.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def rewrite_fixture(self):
        ser_xml, expected_xml = snippet_seq('rewrite-ser')[2:4]

        chart_data = XyChartData()
        series_data = chart_data.add_series('Series 1')
        series_data.add_data_point(1, 2)
        series_data.add_data_point(3, 4)

        rewriter = _XySeriesXmlRewriter(chart_data)
        ser = parse_xml(ser_xml)
        return rewriter, ser, series_data, expected_xml


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


def make_category_chart_data(cat_count, cat_type, ser_count):
    """
    Return a |CategoryChartData| instance populated with *cat_count*
    categories of type *cat_type* and *ser_count* series. Values are
    auto-generated.
    """
    category_labels = {
        date: (
            date(2016, 12, 27),
            date(2016, 12, 28),
            date(2016, 12, 29),
        ),
        float: (1.1, 2.2, 3.3, 4.4, 5.5),
        str: ('Foo', 'Bar', 'Baz', 'Boo', 'Far', 'Faz'),
    }[cat_type]
    point_values = count(1.1, 1.1)
    chart_data = CategoryChartData()
    chart_data.categories = category_labels[:cat_count]
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
