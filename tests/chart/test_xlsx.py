# encoding: utf-8

"""
Test suite for pptx.chart.xlsx module
"""

from __future__ import absolute_import, print_function

import pytest

from xlsxwriter import Workbook
from xlsxwriter.worksheet import Worksheet

from pptx.chart.data import (
    BubbleChartData, Categories, CategoryChartData, CategorySeriesData,
    XyChartData
)
from pptx.chart.xlsx import (
    _BaseWorkbookWriter, BubbleWorkbookWriter, CategoryWorkbookWriter,
    XyWorkbookWriter
)
from pptx.compat import BytesIO

from ..unitutil.mock import ANY, call, class_mock, instance_mock, method_mock


class Describe_BaseWorkbookWriter(object):

    def it_can_generate_a_chart_data_Excel_blob(self, xlsx_blob_fixture):
        workbook_writer, xlsx_file_, workbook_, worksheet_, xlsx_blob = (
            xlsx_blob_fixture
        )
        _xlsx_blob = workbook_writer.xlsx_blob

        workbook_writer._open_worksheet.assert_called_once_with(xlsx_file_)
        workbook_writer._populate_worksheet.assert_called_once_with(
            workbook_writer, workbook_, worksheet_
        )
        assert _xlsx_blob is xlsx_blob

    def it_can_open_a_worksheet_in_a_context(self, open_fixture):
        wb_writer, xlsx_file_, workbook_, worksheet_, Workbook_ = open_fixture

        with wb_writer._open_worksheet(xlsx_file_) as (workbook, worksheet):
            Workbook_.assert_called_once_with(xlsx_file_, {'in_memory': True})
            workbook_.add_worksheet.assert_called_once_with()
            assert workbook is workbook_
            assert worksheet is worksheet_
        workbook_.close.assert_called_once_with()

    def it_raises_on_no_override_of_populate(self, populate_fixture):
        workbook_writer = populate_fixture
        with pytest.raises(NotImplementedError):
            workbook_writer._populate_worksheet(None, None)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def open_fixture(self, xlsx_file_, workbook_, worksheet_, Workbook_):
        workbook_writer = _BaseWorkbookWriter(None)
        workbook_.add_worksheet.return_value = worksheet_
        return workbook_writer, xlsx_file_, workbook_, worksheet_, Workbook_

    @pytest.fixture
    def populate_fixture(self):
        workbook_writer = _BaseWorkbookWriter(None)
        return workbook_writer

    @pytest.fixture
    def xlsx_blob_fixture(
            self, request, xlsx_file_, workbook_, worksheet_,
            _populate_worksheet_, _open_worksheet_, BytesIO_):
        workbook_writer = _BaseWorkbookWriter(None)
        xlsx_blob = 'fooblob'
        BytesIO_.return_value = xlsx_file_
        # to make context manager behavior work
        _open_worksheet_.return_value.__enter__.return_value = (
            workbook_, worksheet_
        )
        xlsx_file_.getvalue.return_value = xlsx_blob
        return (
            workbook_writer, xlsx_file_, workbook_, worksheet_, xlsx_blob
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def BytesIO_(self, request):
        return class_mock(request, 'pptx.chart.xlsx.BytesIO')

    @pytest.fixture
    def _open_worksheet_(self, request):
        return method_mock(request, _BaseWorkbookWriter, '_open_worksheet')

    @pytest.fixture
    def _populate_worksheet_(self, request):
        return method_mock(
            request, _BaseWorkbookWriter, '_populate_worksheet',
            autospec=True
        )

    @pytest.fixture
    def Workbook_(self, request, workbook_):
        return class_mock(
            request, 'pptx.chart.xlsx.Workbook', return_value=workbook_
        )

    @pytest.fixture
    def workbook_(self, request):
        return instance_mock(request, Workbook)

    @pytest.fixture
    def worksheet_(self, request):
        return instance_mock(request, Worksheet)

    @pytest.fixture
    def xlsx_file_(self, request):
        return instance_mock(request, BytesIO)


class DescribeCategoryWorkbookWriter(object):

    def it_knows_the_categories_range_ref(self, categories_ref_fixture):
        workbook_writer, expected_value = categories_ref_fixture
        assert workbook_writer.categories_ref == expected_value

    def it_raises_on_cat_ref_on_no_categories(self, cat_ref_raises_fixture):
        workbook_writer = cat_ref_raises_fixture
        with pytest.raises(ValueError):
            workbook_writer.categories_ref

    def it_knows_the_ref_for_a_series_name(self, ser_name_ref_fixture):
        workbook_writer, series_, expected_value = ser_name_ref_fixture
        assert workbook_writer.series_name_ref(series_) == expected_value

    def it_knows_the_values_range_ref(self, values_ref_fixture):
        workbook_writer, series_, expected_value = values_ref_fixture
        assert workbook_writer.values_ref(series_) == expected_value

    def it_can_populate_a_worksheet_with_chart_data(self, populate_fixture):
        workbook_writer, workbook_, worksheet_ = populate_fixture
        workbook_writer._populate_worksheet(workbook_, worksheet_)
        workbook_writer._write_categories.assert_called_once_with(
            workbook_writer, workbook_, worksheet_
        )
        workbook_writer._write_series.assert_called_once_with(
            workbook_writer, workbook_, worksheet_
        )

    def it_calculates_a_column_reference_to_help(self, col_ref_fixture):
        col_num, expected_value = col_ref_fixture
        column_reference = CategoryWorkbookWriter._column_reference(col_num)
        assert column_reference == expected_value

    def it_raises_on_col_number_out_of_range(self, raises_fixture):
        col_num = raises_fixture
        with pytest.raises(ValueError):
            CategoryWorkbookWriter._column_reference(col_num)

    def it_writes_categories_to_help(self, write_cats_fixture):
        workbook_writer, workbook_, worksheet_ = write_cats_fixture[:3]
        number_format, calls = write_cats_fixture[3:]

        workbook_writer._write_categories(workbook_, worksheet_)

        workbook_.add_format.assert_called_once_with({
            'num_format': number_format
        })
        assert workbook_writer._write_cat_column.call_args_list == calls

    def it_writes_a_category_column_to_help(self, write_cat_fixture):
        workbook_writer, worksheet_, col = write_cat_fixture[:3]
        level, num_format_, calls = write_cat_fixture[3:]
        workbook_writer._write_cat_column(worksheet_, col, level, num_format_)
        assert worksheet_.mock_calls == calls

    def it_writes_series_to_help(self, write_sers_fixture):
        workbook_writer, workbook_, worksheet_, calls = write_sers_fixture
        workbook_writer._write_series(workbook_, worksheet_)
        assert worksheet_.mock_calls == calls

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (1, 1, 'Sheet1!$A$2:$A$2'),
        (1, 3, 'Sheet1!$A$2:$A$4'),
        (2, 4, 'Sheet1!$A$2:$B$5'),
        (3, 8, 'Sheet1!$A$2:$C$9'),
    ])
    def categories_ref_fixture(self, request, chart_data_, categories_):
        depth, leaf_count, expected_value = request.param
        workbook_writer = CategoryWorkbookWriter(chart_data_)
        chart_data_.categories = categories_
        categories_.depth, categories_.leaf_count = depth, leaf_count
        return workbook_writer, expected_value

    @pytest.fixture
    def cat_ref_raises_fixture(self, request, chart_data_, categories_):
        workbook_writer = CategoryWorkbookWriter(chart_data_)
        chart_data_.categories = categories_
        categories_.depth = 0
        return workbook_writer

    @pytest.fixture(params=[
        (2, 'B'), (26, 'Z'), (27, 'AA'), (52, 'AZ'), (53, 'BA'), (676, 'YZ'),
        (677, 'ZA'), (702, 'ZZ'), (703, 'AAA'), (728, 'AAZ'), (729, 'ABA'),
        (1378, 'AZZ'), (1379, 'BAA'), (16384, 'XFD'),
    ])
    def col_ref_fixture(self, request):
        column_number, expected_value = request.param
        return column_number, expected_value

    @pytest.fixture
    def populate_fixture(self, workbook_, worksheet_, _write_categories_,
                         _write_series_):
        workbook_writer = CategoryWorkbookWriter(None)
        return workbook_writer, workbook_, worksheet_

    @pytest.fixture(params=[0, -1, 16385, 30433])
    def raises_fixture(self, request):
        column_number = request.param
        return column_number

    @pytest.fixture(params=[
        (1, 0, 'Sheet1!$B$1'),
        (1, 3, 'Sheet1!$E$1'),
        (3, 0, 'Sheet1!$D$1'),
        (3, 3, 'Sheet1!$G$1'),
    ])
    def ser_name_ref_fixture(self, request, series_data_, categories_):
        cat_depth, series_index, expected_value = request.param
        workbook_writer = CategoryWorkbookWriter(None)
        series_data_.categories = categories_
        categories_.depth = cat_depth
        series_data_.index = series_index
        return workbook_writer, series_data_, expected_value

    @pytest.fixture(params=[
        (1, 0, 3, 'Sheet1!$B$2:$B$4'),
        (1, 1, 3, 'Sheet1!$C$2:$C$4'),
        (2, 0, 5, 'Sheet1!$C$2:$C$6'),
        (3, 2, 7, 'Sheet1!$F$2:$F$8'),
    ])
    def values_ref_fixture(self, request, series_data_, categories_):
        cat_depth, ser_idx, val_count, expected_value = request.param
        workbook_writer = CategoryWorkbookWriter(None)
        series_data_.categories = categories_
        categories_.depth = cat_depth
        series_data_.index = ser_idx
        series_data_.__len__.return_value = val_count
        return workbook_writer, series_data_, expected_value

    @pytest.fixture
    def write_cat_fixture(self, worksheet_):
        workbook_writer = CategoryWorkbookWriter(None)
        col, level = 6, ([0, 'Foo'], [1, 'Bar'])
        num_format = object()
        calls = [
            call.set_column(col, col, 10),
            call.write(1, col, 'Foo', num_format),
            call.write(2, col, 'Bar', num_format),
        ]
        return workbook_writer, worksheet_, col, level, num_format, calls

    @pytest.fixture(params=[
        [
            [[0, 'Foo'], [1, 'Bar'], [2, 'Baz']]
        ],
        [
            [[0, 'CA'], [1, 'NV'], [2, 'NY'], [3, 'NJ']],
            [[0, 'WEST'], [2, 'EAST']]
        ],
    ])
    def write_cats_fixture(self, request, chart_data_, workbook_,
                           worksheet_, categories_, _write_cat_column_):
        levels = request.param
        workbook_writer = CategoryWorkbookWriter(chart_data_)
        number_format = categories_.number_format = '$#0.#'
        num_format = workbook_.add_format.return_value
        _depth = len(levels)
        calls = [
            call(workbook_writer, worksheet_, _depth-idx-1, level, num_format)
            for (idx, level) in enumerate(levels)
        ]
        chart_data_.categories = categories_
        categories_.depth = len(levels)
        categories_.levels = levels
        return workbook_writer, workbook_, worksheet_, number_format, calls

    @pytest.fixture
    def write_sers_fixture(self, request, chart_data_, workbook_, worksheet_,
                           categories_):
        workbook_writer = CategoryWorkbookWriter(chart_data_)
        num_format = workbook_.add_format.return_value
        calls = [
            call.write(0, 1, 'S1'),
            call.write_column(1, 1, (42, 24), num_format),
        ]
        ser = instance_mock(request, CategorySeriesData, values=(42, 24))
        ser.name = 'S1'
        chart_data_.categories = categories_
        categories_.depth = 1
        chart_data_.__iter__.return_value = iter([ser])
        return workbook_writer, workbook_, worksheet_, calls

    # fixture components ---------------------------------------------

    @pytest.fixture
    def categories_(self, request):
        return instance_mock(request, Categories)

    @pytest.fixture
    def chart_data_(self, request):
        return instance_mock(request, CategoryChartData)

    @pytest.fixture
    def series_data_(self, request):
        return instance_mock(request, CategorySeriesData)

    @pytest.fixture
    def workbook_(self, request):
        return instance_mock(request, Workbook)

    @pytest.fixture
    def worksheet_(self, request):
        return instance_mock(request, Worksheet)

    @pytest.fixture
    def _write_cat_column_(self, request):
        return method_mock(
            request, CategoryWorkbookWriter, '_write_cat_column',
            autospec=True
        )

    @pytest.fixture
    def _write_categories_(self, request):
        return method_mock(
            request, CategoryWorkbookWriter, '_write_categories',
            autospec=True
        )

    @pytest.fixture
    def _write_series_(self, request):
        return method_mock(
            request, CategoryWorkbookWriter, '_write_series',
            autospec=True
        )


class DescribeBubbleWorkbookWriter(object):

    def it_can_populate_a_worksheet_with_chart_data(self, populate_fixture):
        workbook_writer, workbook_, worksheet_, expected_calls = (
            populate_fixture
        )
        workbook_writer._populate_worksheet(workbook_, worksheet_)
        assert worksheet_.mock_calls == expected_calls

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def populate_fixture(self, workbook_, worksheet_):
        chart_data = BubbleChartData()
        series_1 = chart_data.add_series('Series 1')
        for pt in ((1, 1.1, 10), (2, 2.2, 20)):
            series_1.add_data_point(*pt)
        series_2 = chart_data.add_series('Series 2')
        for pt in ((3, 3.3, 30), (4, 4.4, 40)):
            series_2.add_data_point(*pt)

        workbook_writer = BubbleWorkbookWriter(chart_data)

        expected_calls = [
            call.write_column(1, 0, [1, 2], ANY),
            call.write(0, 1, 'Series 1'),
            call.write_column(1, 1, [1.1, 2.2], ANY),
            call.write(0, 2, 'Size'),
            call.write_column(1, 2, [10, 20], ANY),

            call.write_column(5, 0, [3, 4], ANY),
            call.write(4, 1, 'Series 2'),
            call.write_column(5, 1, [3.3, 4.4], ANY),
            call.write(4, 2, 'Size'),
            call.write_column(5, 2, [30, 40], ANY),
        ]
        return workbook_writer, workbook_, worksheet_, expected_calls

    # fixture components ---------------------------------------------

    @pytest.fixture
    def workbook_(self, request):
        return instance_mock(request, Workbook)

    @pytest.fixture
    def worksheet_(self, request):
        return instance_mock(request, Worksheet)


class DescribeXyWorkbookWriter(object):

    def it_can_generate_a_chart_data_Excel_blob(self, xlsx_blob_fixture):
        workbook_writer, _open_worksheet_, xlsx_file_ = xlsx_blob_fixture[:3]
        _populate_worksheet_, workbook_, worksheet_ = xlsx_blob_fixture[3:6]
        xlsx_blob_ = xlsx_blob_fixture[6]

        xlsx_blob = workbook_writer.xlsx_blob

        _open_worksheet_.assert_called_once_with(xlsx_file_)
        _populate_worksheet_.assert_called_once_with(workbook_, worksheet_)
        assert xlsx_blob is xlsx_blob_

    def it_can_populate_a_worksheet_with_chart_data(self, populate_fixture):
        workbook_writer, workbook_, worksheet_, expected_calls = (
            populate_fixture
        )
        workbook_writer._populate_worksheet(workbook_, worksheet_)
        assert worksheet_.mock_calls == expected_calls

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def populate_fixture(self, workbook_, worksheet_):
        chart_data = XyChartData()
        series_1 = chart_data.add_series('Series 1')
        for pt in ((1, 1.1), (2, 2.2)):
            series_1.add_data_point(*pt)
        series_2 = chart_data.add_series('Series 2')
        for pt in ((3, 3.3), (4, 4.4)):
            series_2.add_data_point(*pt)

        workbook_writer = XyWorkbookWriter(chart_data)

        expected_calls = [
            call.write_column(1, 0, [1, 2], ANY),
            call.write(0, 1, 'Series 1'),
            call.write_column(1, 1, [1.1, 2.2], ANY),

            call.write_column(5, 0, [3, 4], ANY),
            call.write(4, 1, 'Series 2'),
            call.write_column(5, 1, [3.3, 4.4], ANY)
        ]
        return workbook_writer, workbook_, worksheet_, expected_calls

    @pytest.fixture
    def xlsx_blob_fixture(
            self, request, xlsx_file_, BytesIO_, _open_worksheet_, workbook_,
            worksheet_, _populate_worksheet_, xlsx_blob_):
        workbook_writer = XyWorkbookWriter(None)
        return (
            workbook_writer, _open_worksheet_, xlsx_file_,
            _populate_worksheet_, workbook_, worksheet_, xlsx_blob_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def BytesIO_(self, request, xlsx_file_):
        return class_mock(
            request, 'pptx.chart.xlsx.BytesIO', return_value=xlsx_file_
        )

    @pytest.fixture
    def _open_worksheet_(self, request, workbook_, worksheet_):
        open_worksheet_ = method_mock(
            request, XyWorkbookWriter, '_open_worksheet'
        )
        # to make context manager behavior work
        open_worksheet_.return_value.__enter__.return_value = (
            workbook_, worksheet_
        )
        return open_worksheet_

    @pytest.fixture
    def _populate_worksheet_(self, request):
        return method_mock(request, XyWorkbookWriter, '_populate_worksheet')

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
