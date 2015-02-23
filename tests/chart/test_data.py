# encoding: utf-8

"""
Test suite for pptx.chart.data module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest


from pptx.chart.data import ChartData, _SeriesData
from pptx.enum.base import EnumValue

from ..unitutil.mock import class_mock, instance_mock


class DescribeChartData(object):

    def it_knows_its_current_categories(self, categories_get_fixture):
        chart_data, expected_value = categories_get_fixture
        assert chart_data.categories == expected_value

    def it_can_change_its_categories(self, categories_set_fixture):
        chart_data, new_value, expected_value = categories_set_fixture
        chart_data.categories = new_value
        assert chart_data.categories == expected_value

    def it_provides_access_to_its_current_series_data(self, series_fixture):
        chart_data, expected_value = series_fixture
        assert chart_data.series == expected_value

    def it_can_add_a_series(self, add_series_fixture):
        chart_data, name, values, _SeriesData_, series_data_ = (
            add_series_fixture
        )
        chart_data.add_series(name, values)
        _SeriesData_.assert_called_once_with(
            0, name, values, chart_data._categories, 0
        )
        assert chart_data._series_lst[0] is series_data_

    def it_can_generate_chart_part_XML_for_its_data(self, xml_bytes_fixture):
        chart_data, chart_type_, ChartXmlWriter_ = xml_bytes_fixture[:3]
        expected_bytes, series_lst_ = xml_bytes_fixture[3:]

        xml_bytes = chart_data.xml_bytes(chart_type_)

        ChartXmlWriter_.assert_called_once_with(chart_type_, series_lst_)
        assert xml_bytes == expected_bytes

    def it_can_provide_its_data_as_an_Excel_workbook(self, xlsx_fixture):
        chart_data, WorkbookWriter_ = xlsx_fixture[:2]
        categories, series_, xlsx_blob_ = xlsx_fixture[2:]
        xlsx_blob = chart_data.xlsx_blob
        WorkbookWriter_.xlsx_blob.assert_called_once_with(
            categories, series_
        )
        assert xlsx_blob is xlsx_blob_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_series_fixture(
            self, request, categories, _SeriesData_, series_data_):
        chart_data = ChartData()
        chart_data.categories = categories
        name = 'Series Foo'
        values = (1.1, 2.2, 3.3)
        return chart_data, name, values, _SeriesData_, series_data_

    @pytest.fixture
    def categories_get_fixture(self, categories):
        chart_data = ChartData()
        chart_data._categories = list(categories)
        expected_value = categories
        return chart_data, expected_value

    @pytest.fixture(params=[
        (['Foo', 'Bar'],       ('Foo', 'Bar')),
        (iter(['Foo', 'Bar']), ('Foo', 'Bar')),
    ])
    def categories_set_fixture(self, request):
        new_value, expected_value = request.param
        chart_data = ChartData()
        return chart_data, new_value, expected_value

    @pytest.fixture
    def series_fixture(self, series_data_):
        chart_data = ChartData()
        chart_data._series_lst = [series_data_, series_data_]
        expected_value = (series_data_, series_data_)
        return chart_data, expected_value

    @pytest.fixture
    def xlsx_fixture(
            self, request, WorkbookWriter_, categories, series_lst_,
            xlsx_blob_):
        chart_data = ChartData()
        chart_data.categories = categories
        chart_data._series_lst = series_lst_
        return (
            chart_data, WorkbookWriter_, categories, series_lst_, xlsx_blob_
        )

    @pytest.fixture
    def xml_bytes_fixture(self, chart_type_, ChartXmlWriter_, series_lst_):
        chart_data = ChartData()
        chart_data._series_lst = series_lst_
        expected_bytes = 'ƒøØßår'.encode('utf-8')
        return (
            chart_data, chart_type_, ChartXmlWriter_, expected_bytes,
            series_lst_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def categories(self):
        return ('Foo', 'Bar', 'Baz')

    @pytest.fixture
    def ChartXmlWriter_(self, request):
        ChartXmlWriter_ = class_mock(
            request, 'pptx.chart.data.ChartXmlWriter'
        )
        ChartXmlWriter_.return_value.xml = 'ƒøØßår'
        return ChartXmlWriter_

    @pytest.fixture
    def chart_type_(self, request):
        return instance_mock(request, EnumValue)

    @pytest.fixture
    def _SeriesData_(self, request, series_data_):
        return class_mock(
            request, 'pptx.chart.data._SeriesData', return_value=series_data_
        )

    @pytest.fixture
    def series_data_(self, request):
        return instance_mock(request, _SeriesData)

    @pytest.fixture
    def series_lst_(self, request):
        return instance_mock(request, list)

    @pytest.fixture
    def WorkbookWriter_(self, request, xlsx_blob_):
        WorkbookWriter_ = class_mock(
            request, 'pptx.chart.data.WorkbookWriter'
        )
        WorkbookWriter_.xlsx_blob.return_value = xlsx_blob_
        return WorkbookWriter_

    @pytest.fixture
    def xlsx_blob_(self, request):
        return instance_mock(request, bytes)
