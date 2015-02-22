# encoding: utf-8

"""
Test suite for pptx.chart.xlsx module
"""

from __future__ import absolute_import, print_function

import pytest

from zipfile import ZipFile

from xlsxwriter import Workbook
from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet

from pptx.chart.data import _SeriesData
from pptx.chart.xlsx import WorkbookWriter
from pptx.compat import BytesIO

from ..unitutil.mock import call, class_mock, instance_mock, method_mock


class Describe_WorkbookWriter(object):

    def it_can_generate_a_chart_data_Excel_blob(self, xlsx_blob_fixture):
        categories_, series_, xlsx_file_ = xlsx_blob_fixture[:3]
        _populate_worksheet_, workbook_, worksheet_ = xlsx_blob_fixture[3:6]
        xlsx_blob_ = xlsx_blob_fixture[6]

        xlsx_blob = WorkbookWriter.xlsx_blob(categories_, series_)

        WorkbookWriter._open_worksheet.assert_called_once_with(xlsx_file_)
        _populate_worksheet_.assert_called_once_with(
            workbook_, worksheet_, categories_, series_
        )
        assert xlsx_blob is xlsx_blob_

    def it_can_open_a_worksheet_in_a_context(self):
        xlsx_file = BytesIO()
        with WorkbookWriter._open_worksheet(xlsx_file) as (wrkbook, wrksht):
            assert isinstance(wrkbook, Workbook)
            assert isinstance(wrksht,  Worksheet)
        zipf = ZipFile(xlsx_file)
        assert 'xl/worksheets/sheet1.xml' in zipf.namelist()
        zipf.close

    def it_can_populate_a_worksheet_with_chart_data(self, populate_fixture):
        workbook_, worksheet_, categories = populate_fixture[:3]
        series, expected_calls = populate_fixture[3:]
        WorkbookWriter._populate_worksheet(
            workbook_, worksheet_, categories, series
        )
        assert worksheet_.mock_calls == expected_calls

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def populate_fixture(self, workbook_, worksheet_, format_):
        categories = ('Foo', 'Bar')
        series = (
            _SeriesData(0, 'Series 1', (1.1, 2.2), categories, 0),
            _SeriesData(1, 'Series 2', (3.3, 4.4), categories, 0)
        )
        expected_calls = [
            call.write_column(1, 0, ('Foo', 'Bar')),
            call.write(0, 1, 'Series 1'),
            call.write_column(1, 1, (1.1, 2.2), format_),
            call.write(0, 2, 'Series 2'),
            call.write_column(1, 2, (3.3, 4.4), format_)
        ]
        workbook_.add_format.return_value = format_
        return workbook_, worksheet_, categories, series, expected_calls

    @pytest.fixture
    def xlsx_blob_fixture(
            self, request, categories, series_lst_, xlsx_file_, BytesIO_,
            _open_worksheet_, workbook_, worksheet_, _populate_worksheet_,
            xlsx_blob_):
        return (
            categories, series_lst_, xlsx_file_, _populate_worksheet_,
            workbook_, worksheet_, xlsx_blob_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def BytesIO_(self, request, xlsx_file_):
        return class_mock(
            request, 'pptx.chart.xlsx.BytesIO', return_value=xlsx_file_
        )

    @pytest.fixture
    def categories(self):
        return ('Foo', 'Bar')

    @pytest.fixture
    def format_(self, request):
        return instance_mock(request, Format)

    @pytest.fixture
    def _open_worksheet_(self, request, workbook_, worksheet_):
        open_worksheet_ = method_mock(
            request, WorkbookWriter, '_open_worksheet'
        )
        # to make context manager behavior work
        open_worksheet_.return_value.__enter__.return_value = (
            workbook_, worksheet_
        )
        return open_worksheet_

    @pytest.fixture
    def _populate_worksheet_(self, request):
        return method_mock(request, WorkbookWriter, '_populate_worksheet')

    @pytest.fixture
    def series_lst_(self, request):
        return instance_mock(request, list)

    @pytest.fixture
    def workbook_(self, request):
        return instance_mock(request, Workbook)

    @pytest.fixture
    def worksheet_(self, request):
        return instance_mock(request, Worksheet)

    @pytest.fixture
    def xlsx_blob_(self, request):
        return instance_mock(request, bytes)

    @pytest.fixture
    def xlsx_file_(self, request, xlsx_blob_):
        xlsx_file_ = instance_mock(request, BytesIO)
        xlsx_file_.getvalue.return_value = xlsx_blob_
        return xlsx_file_
