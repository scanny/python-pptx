# encoding: utf-8

"""
Test suite for pptx.chart.xmlwriter module
"""

from __future__ import absolute_import, print_function, unicode_literals

from itertools import islice

import pytest


from pptx.chart.data import ChartData
from pptx.chart.xmlwriter import (
    _BarChartXmlWriter, ChartXmlWriter, _LineChartXmlWriter,
    _PieChartXmlWriter
)
from pptx.enum.chart import XL_CHART_TYPE

from ..unitutil import count
from ..unitutil.file import snippet_text
from ..unitutil.mock import class_mock, instance_mock


class DescribeChartXmlWriter(object):

    def it_contructs_an_xml_writer_for_a_chart_type(self, call_fixture):
        chart_type, series_seq_, XmlWriterClass_, xml_writer_ = call_fixture
        xml_writer = ChartXmlWriter(chart_type, series_seq_)
        XmlWriterClass_.assert_called_once_with(chart_type, series_seq_)
        assert xml_writer is xml_writer_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('BAR_CLUSTERED',    _BarChartXmlWriter),
        ('BAR_STACKED_100',  _BarChartXmlWriter),
        ('COLUMN_CLUSTERED', _BarChartXmlWriter),
        ('LINE',             _LineChartXmlWriter),
        ('PIE',              _PieChartXmlWriter),
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


class Describe_BarChartXmlWriter(object):

    def it_can_generate_xml_for_bar_type_charts(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('BAR_CLUSTERED',    2, 2, '2x2-bar-clustered'),
        ('BAR_STACKED_100',  2, 2, '2x2-bar-stacked-100'),
        ('COLUMN_CLUSTERED', 2, 2, '2x2-column-clustered'),
    ])
    def xml_fixture(self, request):
        enum_member, cat_count, ser_count, snippet_name = request.param
        chart_type = getattr(XL_CHART_TYPE, enum_member)
        series_data_seq = make_series_data_seq(cat_count, ser_count)
        xml_writer = _BarChartXmlWriter(chart_type, series_data_seq)
        expected_xml = snippet_text(snippet_name)
        return xml_writer, expected_xml


class Describe_LineChartXmlWriter(object):

    def it_can_generate_xml_for_a_line_chart(self, xml_fixture):
        xml_writer, expected_xml = xml_fixture
        assert xml_writer.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def xml_fixture(self, request):
        series_data_seq = make_series_data_seq(cat_count=2, ser_count=2)
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
        series_data_seq = make_series_data_seq(cat_count=3, ser_count=1)
        xml_writer = _PieChartXmlWriter(XL_CHART_TYPE.PIE, series_data_seq)
        expected_xml = snippet_text('3x1-pie')
        return xml_writer, expected_xml


# helpers ------------------------------------------------------------

def make_series_data_seq(cat_count, ser_count):
    """
    Return a sequence of |_SeriesData| objects populated with *cat_count*
    category names and *ser_count* sequences of point values. Values are
    auto-generated.
    """
    category_names = ('Foo', 'Bar', 'Baz', 'Boo', 'Far', 'Faz')
    point_values = count(1.1, 1.1)
    chart_data = ChartData()
    chart_data.categories = category_names[:cat_count]
    for idx in range(ser_count):
        series_title = 'Series %d' % (idx+1)
        series_values = tuple(islice(point_values, cat_count))
        series_values = [round(x*10)/10.0 for x in series_values]
        chart_data.add_series(series_title, series_values)
    return chart_data.series
