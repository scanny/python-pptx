# encoding: utf-8

"""
Test suite for pptx.chart.xmlwriter module
"""

from __future__ import absolute_import, print_function

import pytest


from pptx.chart.xmlwriter import (
    _BarChartXmlWriter, ChartXmlWriter, _LineChartXmlWriter,
    _PieChartXmlWriter
)
from pptx.enum.chart import XL_CHART_TYPE

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
