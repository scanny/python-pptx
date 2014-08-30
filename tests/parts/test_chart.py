# encoding: utf-8

"""
Test suite for pptx.parts.chart module.
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.chart import Chart
from pptx.chart.data import ChartData
from pptx.enum.base import EnumValue
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import OpcPackage
from pptx.opc.packuri import PackURI
from pptx.oxml.chart.chart import CT_ChartSpace
from pptx.parts.chart import ChartPart, ChartWorkbook

from ..unitutil.mock import class_mock, instance_mock, method_mock


class DescribeChartPart(object):

    def it_can_construct_from_chart_type_and_data(self, new_fixture):
        chart_type_, chart_data_, package_ = new_fixture[:3]
        partname_template, load_, partname_ = new_fixture[3:6]
        content_type, chart_blob_, chart_part_, xlsx_blob_ = new_fixture[6:]

        chart_part = ChartPart.new(chart_type_, chart_data_, package_)

        chart_data_.xml_bytes.assert_called_once_with(chart_type_)
        package_.next_partname.assert_called_once_with(partname_template)
        load_.assert_called_once_with(
            partname_, content_type, chart_blob_, package_
        )
        chart_workbook_ = chart_part_.chart_workbook
        chart_workbook_.update_from_xlsx_blob.assert_called_once_with(
            xlsx_blob_
        )
        assert chart_part is chart_part_

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

    @pytest.fixture
    def new_fixture(
            self, chart_type_, chart_data_, package_, load_, partname_,
            chart_blob_, chart_part_, xlsx_blob_):
        partname_template = '/ppt/charts/chart%d.xml'
        content_type = CT.DML_CHART
        return (
            chart_type_, chart_data_, package_, partname_template, load_,
            partname_, content_type, chart_blob_, chart_part_, xlsx_blob_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Chart_(self, request, chart_):
        return class_mock(
            request, 'pptx.parts.chart.Chart', return_value=chart_
        )

    @pytest.fixture
    def chartSpace_(self, request):
        return instance_mock(request, CT_ChartSpace)

    @pytest.fixture
    def chart_(self, request):
        return instance_mock(request, Chart)

    @pytest.fixture
    def chart_blob_(self, request):
        return instance_mock(request, bytes)

    @pytest.fixture
    def chart_data_(self, request, chart_blob_, xlsx_blob_):
        chart_data_ = instance_mock(request, ChartData)
        chart_data_.xml_bytes.return_value = chart_blob_
        chart_data_.xlsx_blob = xlsx_blob_
        return chart_data_

    @pytest.fixture
    def chart_part_(self, request, chart_workbook_):
        chart_part_ = instance_mock(request, ChartPart)
        chart_part_.chart_workbook = chart_workbook_
        return chart_part_

    @pytest.fixture
    def chart_type_(self, request):
        return instance_mock(request, EnumValue)

    @pytest.fixture
    def chart_workbook_(self, request):
        return instance_mock(request, ChartWorkbook)

    @pytest.fixture
    def load_(self, request, chart_part_):
        return method_mock(
            request, ChartPart, 'load', return_value=chart_part_
        )

    @pytest.fixture
    def package_(self, request, partname_):
        package_ = instance_mock(request, OpcPackage)
        package_.next_partname.return_value = partname_
        return package_

    @pytest.fixture
    def partname_(self, request):
        return instance_mock(request, PackURI)

    @pytest.fixture
    def xlsx_blob_(self, request):
        return instance_mock(request, bytes)
