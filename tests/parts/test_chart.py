# encoding: utf-8

"""
Test suite for pptx.parts.chart module.
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.chart import Chart
from pptx.chart.data import ChartData
from pptx.enum.base import EnumValue
from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.opc.package import OpcPackage
from pptx.opc.packuri import PackURI
from pptx.oxml.chart.chart import CT_ChartSpace
from pptx.parts.chart import ChartPart, ChartWorkbook
from pptx.parts.embeddedpackage import EmbeddedXlsxPart

from ..unitutil.cxml import element, xml
from ..unitutil.mock import (
    class_mock, instance_mock, method_mock, property_mock
)


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

    def it_provides_access_to_the_chart_workbook(self, workbook_fixture):
        chart_part, ChartWorkbook_, chartSpace_, chart_workbook_ = (
            workbook_fixture
        )
        chart_workbook = chart_part.chart_workbook
        ChartWorkbook_.assert_called_once_with(chartSpace_, chart_part)
        assert chart_workbook is chart_workbook_

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

    @pytest.fixture
    def workbook_fixture(self, chartSpace_, ChartWorkbook_, chart_workbook_):
        chart_part = ChartPart(None, None, chartSpace_)
        return chart_part, ChartWorkbook_, chartSpace_, chart_workbook_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ChartWorkbook_(self, request, chart_workbook_):
        return class_mock(
            request, 'pptx.parts.chart.ChartWorkbook',
            return_value=chart_workbook_
        )

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


class DescribeChartWorkbook(object):

    def it_can_get_the_chart_xlsx_part(self, xlsx_part_get_fixture):
        chart_data, expected_object = xlsx_part_get_fixture
        assert chart_data.xlsx_part is expected_object

    def it_can_change_the_chart_xlsx_part(self, xlsx_part_set_fixture):
        chart_data, xlsx_part_, expected_xml = xlsx_part_set_fixture
        chart_data.xlsx_part = xlsx_part_
        chart_data._chart_part.relate_to.assert_called_once_with(
            xlsx_part_, RT.PACKAGE
        )
        assert chart_data._chartSpace.xml == expected_xml

    def it_adds_an_xlsx_part_on_update_if_needed(self, add_part_fixture):
        chart_data, xlsx_blob_, EmbeddedXlsxPart_ = add_part_fixture[:3]
        package_, xlsx_part_prop_, xlsx_part_ = add_part_fixture[3:]

        chart_data.update_from_xlsx_blob(xlsx_blob_)

        EmbeddedXlsxPart_.new.assert_called_once_with(xlsx_blob_, package_)
        xlsx_part_prop_.assert_called_with(xlsx_part_)

    def but_replaces_xlsx_blob_when_part_exists(self, update_blob_fixture):
        chart_data, xlsx_blob_ = update_blob_fixture
        chart_data.update_from_xlsx_blob(xlsx_blob_)
        assert chart_data.xlsx_part.blob is xlsx_blob_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_part_fixture(
            self, request, chart_part_, xlsx_blob_, EmbeddedXlsxPart_,
            package_, xlsx_part_, xlsx_part_prop_):
        chartSpace_cxml = 'c:chartSpace'
        chart_data = ChartWorkbook(element(chartSpace_cxml), chart_part_)
        xlsx_part_prop_.return_value = None
        return (
            chart_data, xlsx_blob_, EmbeddedXlsxPart_, package_,
            xlsx_part_prop_, xlsx_part_
        )

    @pytest.fixture
    def update_blob_fixture(self, request, xlsx_blob_, xlsx_part_prop_):
        chart_data = ChartWorkbook(None, None)
        return chart_data, xlsx_blob_

    @pytest.fixture(params=[
        ('c:chartSpace', None),
        ('c:chartSpace/c:externalData{r:id=rId42}', 'rId42'),
    ])
    def xlsx_part_get_fixture(self, request, chart_part_, xlsx_part_):
        chartSpace_cxml, xlsx_part_rId = request.param
        chart_data = ChartWorkbook(element(chartSpace_cxml), chart_part_)
        expected_object = xlsx_part_ if xlsx_part_rId else None
        return chart_data, expected_object

    @pytest.fixture(params=[
        ('c:chartSpace{r:a=b}', 'c:chartSpace{r:a=b}/c:externalData{r:id=rId'
         '42}/c:autoUpdate{val=0}'),
        ('c:chartSpace/c:externalData{r:id=rId66}',
         'c:chartSpace/c:externalData{r:id=rId42}'),
    ])
    def xlsx_part_set_fixture(self, request, chart_part_, xlsx_part_):
        chartSpace_cxml, expected_cxml = request.param
        chart_data = ChartWorkbook(element(chartSpace_cxml), chart_part_)
        expected_xml = xml(expected_cxml)
        return chart_data, xlsx_part_, expected_xml

    # fixture components ---------------------------------------------

    @pytest.fixture
    def chart_part_(self, request, package_, xlsx_part_):
        chart_part_ = instance_mock(request, ChartPart)
        chart_part_.package = package_
        chart_part_.related_parts = {'rId42': xlsx_part_}
        chart_part_.relate_to.return_value = 'rId42'
        return chart_part_

    @pytest.fixture
    def EmbeddedXlsxPart_(self, request, xlsx_part_):
        EmbeddedXlsxPart_ = class_mock(
            request, 'pptx.parts.chart.EmbeddedXlsxPart'
        )
        EmbeddedXlsxPart_.new.return_value = xlsx_part_
        return EmbeddedXlsxPart_

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, OpcPackage)

    @pytest.fixture
    def xlsx_blob_(self, request):
        return instance_mock(request, bytes)

    @pytest.fixture
    def xlsx_part_(self, request):
        return instance_mock(request, EmbeddedXlsxPart)

    @pytest.fixture
    def xlsx_part_prop_(self, request, xlsx_part_):
        return property_mock(request, ChartWorkbook, 'xlsx_part')
