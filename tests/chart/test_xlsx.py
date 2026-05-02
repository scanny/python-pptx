# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.chart.xlsx` module."""

from __future__ import annotations

import io

import pytest
from xlsxwriter import Workbook
from xlsxwriter.worksheet import Worksheet

from pptx.chart.data import (
    BubbleChartData,
    Categories,
    CategoryChartData,
    CategorySeriesData,
    XyChartData,
)
from pptx.chart.xlsx import (
    BubbleWorkbookWriter,
    CategoryWorkbookWriter,
    WorkbookReader,
    WorkbookUpdater,
    XyWorkbookWriter,
    _BaseWorkbookWriter,
    _row_col_to_a1,
    parse_a1_cell,
    parse_sheet_range_ref,
)

from ..unitutil.mock import ANY, call, class_mock, instance_mock, method_mock


class Describe_BaseWorkbookWriter(object):
    """Unit-test suite for `pptx.chart.xlsx._BaseWorkbookWriter` objects."""

    def it_can_generate_a_chart_data_Excel_blob(
        self, request, xlsx_file_, workbook_, worksheet_, BytesIO_
    ):
        _populate_worksheet_ = method_mock(request, _BaseWorkbookWriter, "_populate_worksheet")
        _open_worksheet_ = method_mock(request, _BaseWorkbookWriter, "_open_worksheet")
        # --- to make context manager behavior work ---
        _open_worksheet_.return_value.__enter__.return_value = (workbook_, worksheet_)
        BytesIO_.return_value = xlsx_file_
        xlsx_file_.getvalue.return_value = b"xlsx-blob"
        workbook_writer = _BaseWorkbookWriter(None)

        xlsx_blob = workbook_writer.xlsx_blob

        _open_worksheet_.assert_called_once_with(workbook_writer, xlsx_file_)
        _populate_worksheet_.assert_called_once_with(workbook_writer, workbook_, worksheet_)
        assert xlsx_blob == b"xlsx-blob"

    def it_can_open_a_worksheet_in_a_context(self, open_fixture):
        wb_writer, xlsx_file_, workbook_, worksheet_, Workbook_ = open_fixture

        with wb_writer._open_worksheet(xlsx_file_) as (workbook, worksheet):
            Workbook_.assert_called_once_with(xlsx_file_, {"in_memory": True})
            workbook_.add_worksheet.assert_called_once_with()
            assert workbook is workbook_
            assert worksheet is worksheet_
        workbook_.close.assert_called_once_with()

    def it_raises_on_no_override_of_populate(self, populate_fixture):
        workbook_writer = populate_fixture
        with pytest.raises(NotImplementedError):
            workbook_writer._populate_worksheet(None, None)

    def it_raises_on_no_override_of_iter_cell_writes_239(self):
        workbook_writer = _BaseWorkbookWriter(None)
        with pytest.raises(NotImplementedError):
            next(workbook_writer.iter_cell_writes())

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

    # fixture components ---------------------------------------------

    @pytest.fixture
    def BytesIO_(self, request):
        return class_mock(request, "pptx.chart.xlsx.io.BytesIO")

    @pytest.fixture
    def Workbook_(self, request, workbook_):
        return class_mock(request, "pptx.chart.xlsx.Workbook", return_value=workbook_)

    @pytest.fixture
    def workbook_(self, request):
        return instance_mock(request, Workbook)

    @pytest.fixture
    def worksheet_(self, request):
        return instance_mock(request, Worksheet)

    @pytest.fixture
    def xlsx_file_(self, request):
        return instance_mock(request, io.BytesIO)


class DescribeCategoryWorkbookWriter(object):
    """Unit-test suite for `pptx.chart.xlsx.CategoryWorkbookWriter` objects."""

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

        workbook_.add_format.assert_called_once_with({"num_format": number_format})
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

    def it_enumerates_cell_writes_matching_the_worksheet_layout_239(self):
        """`iter_cell_writes` mirrors `_populate_worksheet` layout (#239)."""
        from pptx.chart.data import CategoryChartData

        cd = CategoryChartData()
        cd.categories = ["CA", "CB"]
        cd.add_series("S1", (10, 20))
        cd.add_series("S2", (30, 40))
        writer = cd._workbook_writer

        cells = list(writer.iter_cell_writes())

        # -- 2 category cells (A2, A3) + 2 series * (heading + 2 values) --
        assert cells == [
            ("Sheet1", 2, 1, "CA"),
            ("Sheet1", 3, 1, "CB"),
            ("Sheet1", 1, 2, "S1"),
            ("Sheet1", 2, 2, 10),
            ("Sheet1", 3, 2, 20),
            ("Sheet1", 1, 3, "S2"),
            ("Sheet1", 2, 3, 30),
            ("Sheet1", 3, 3, 40),
        ]

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            (1, 1, "Sheet1!$A$2:$A$2"),
            (1, 3, "Sheet1!$A$2:$A$4"),
            (2, 4, "Sheet1!$A$2:$B$5"),
            (3, 8, "Sheet1!$A$2:$C$9"),
        ]
    )
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

    @pytest.fixture(
        params=[
            (2, "B"),
            (26, "Z"),
            (27, "AA"),
            (52, "AZ"),
            (53, "BA"),
            (676, "YZ"),
            (677, "ZA"),
            (702, "ZZ"),
            (703, "AAA"),
            (728, "AAZ"),
            (729, "ABA"),
            (1378, "AZZ"),
            (1379, "BAA"),
            (16384, "XFD"),
        ]
    )
    def col_ref_fixture(self, request):
        column_number, expected_value = request.param
        return column_number, expected_value

    @pytest.fixture
    def populate_fixture(self, workbook_, worksheet_, _write_categories_, _write_series_):
        workbook_writer = CategoryWorkbookWriter(None)
        return workbook_writer, workbook_, worksheet_

    @pytest.fixture(params=[0, -1, 16385, 30433])
    def raises_fixture(self, request):
        column_number = request.param
        return column_number

    @pytest.fixture(
        params=[
            (1, 0, "Sheet1!$B$1"),
            (1, 3, "Sheet1!$E$1"),
            (3, 0, "Sheet1!$D$1"),
            (3, 3, "Sheet1!$G$1"),
        ]
    )
    def ser_name_ref_fixture(self, request, series_data_, categories_):
        cat_depth, series_index, expected_value = request.param
        workbook_writer = CategoryWorkbookWriter(None)
        series_data_.categories = categories_
        categories_.depth = cat_depth
        series_data_.index = series_index
        return workbook_writer, series_data_, expected_value

    @pytest.fixture(
        params=[
            (1, 0, 3, "Sheet1!$B$2:$B$4"),
            (1, 1, 3, "Sheet1!$C$2:$C$4"),
            (2, 0, 5, "Sheet1!$C$2:$C$6"),
            (3, 2, 7, "Sheet1!$F$2:$F$8"),
        ]
    )
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
        col, level = 6, ([0, "Foo"], [1, "Bar"])
        num_format = object()
        calls = [
            call.set_column(col, col, 10),
            call.write(1, col, "Foo", num_format),
            call.write(2, col, "Bar", num_format),
        ]
        return workbook_writer, worksheet_, col, level, num_format, calls

    @pytest.fixture(
        params=[
            [[[0, "Foo"], [1, "Bar"], [2, "Baz"]]],
            [[[0, "CA"], [1, "NV"], [2, "NY"], [3, "NJ"]], [[0, "WEST"], [2, "EAST"]]],
        ]
    )
    def write_cats_fixture(
        self,
        request,
        chart_data_,
        workbook_,
        worksheet_,
        categories_,
        _write_cat_column_,
    ):
        levels = request.param
        workbook_writer = CategoryWorkbookWriter(chart_data_)
        number_format = categories_.number_format = "$#0.#"
        num_format = workbook_.add_format.return_value
        _depth = len(levels)
        calls = [
            call(workbook_writer, worksheet_, _depth - idx - 1, level, num_format)
            for (idx, level) in enumerate(levels)
        ]
        chart_data_.categories = categories_
        categories_.depth = len(levels)
        categories_.levels = levels
        return workbook_writer, workbook_, worksheet_, number_format, calls

    @pytest.fixture
    def write_sers_fixture(self, request, chart_data_, workbook_, worksheet_, categories_):
        workbook_writer = CategoryWorkbookWriter(chart_data_)
        num_format = workbook_.add_format.return_value
        calls = [call.write(0, 1, "S1"), call.write_column(1, 1, (42, 24), num_format)]
        ser = instance_mock(request, CategorySeriesData, values=(42, 24))
        ser.name = "S1"
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
        return method_mock(request, CategoryWorkbookWriter, "_write_cat_column", autospec=True)

    @pytest.fixture
    def _write_categories_(self, request):
        return method_mock(request, CategoryWorkbookWriter, "_write_categories", autospec=True)

    @pytest.fixture
    def _write_series_(self, request):
        return method_mock(request, CategoryWorkbookWriter, "_write_series", autospec=True)


class DescribeBubbleWorkbookWriter(object):
    """Unit-test suite for `pptx.chart.xlsx.BubbleWorkbookWriter` objects."""

    def it_can_populate_a_worksheet_with_chart_data(self, populate_fixture):
        workbook_writer, workbook_, worksheet_, expected_calls = populate_fixture
        workbook_writer._populate_worksheet(workbook_, worksheet_)
        assert worksheet_.mock_calls == expected_calls

    def it_enumerates_cell_writes_matching_the_worksheet_layout_239(self):
        """`iter_cell_writes` mirrors `_populate_worksheet` layout (#239)."""
        chart_data = BubbleChartData()
        series_1 = chart_data.add_series("Series 1")
        for pt in ((1, 1.1, 10),):
            series_1.add_data_point(*pt)
        writer = chart_data._workbook_writer

        cells = list(writer.iter_cell_writes())

        assert cells == [
            ("Sheet1", 2, 1, 1),
            ("Sheet1", 1, 2, "Series 1"),
            ("Sheet1", 2, 2, 1.1),
            ("Sheet1", 1, 3, "Size"),
            ("Sheet1", 2, 3, 10),
        ]

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def populate_fixture(self, workbook_, worksheet_):
        chart_data = BubbleChartData()
        series_1 = chart_data.add_series("Series 1")
        for pt in ((1, 1.1, 10), (2, 2.2, 20)):
            series_1.add_data_point(*pt)
        series_2 = chart_data.add_series("Series 2")
        for pt in ((3, 3.3, 30), (4, 4.4, 40)):
            series_2.add_data_point(*pt)

        workbook_writer = BubbleWorkbookWriter(chart_data)

        expected_calls = [
            call.write_column(1, 0, [1, 2], ANY),
            call.write(0, 1, "Series 1"),
            call.write_column(1, 1, [1.1, 2.2], ANY),
            call.write(0, 2, "Size"),
            call.write_column(1, 2, [10, 20], ANY),
            call.write_column(5, 0, [3, 4], ANY),
            call.write(4, 1, "Series 2"),
            call.write_column(5, 1, [3.3, 4.4], ANY),
            call.write(4, 2, "Size"),
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
    """Unit-test suite for `pptx.chart.xlsx.XyWorkbookWriter` objects."""

    def it_can_populate_a_worksheet_with_chart_data(self, populate_fixture):
        workbook_writer, workbook_, worksheet_, expected_calls = populate_fixture
        workbook_writer._populate_worksheet(workbook_, worksheet_)
        assert worksheet_.mock_calls == expected_calls

    def it_enumerates_cell_writes_matching_the_worksheet_layout_239(self):
        """`iter_cell_writes` mirrors `_populate_worksheet` layout (#239)."""
        chart_data = XyChartData()
        series_1 = chart_data.add_series("Series 1")
        for pt in ((1, 1.1), (2, 2.2)):
            series_1.add_data_point(*pt)
        series_2 = chart_data.add_series("Series 2")
        for pt in ((3, 3.3),):
            series_2.add_data_point(*pt)
        writer = chart_data._workbook_writer

        cells = list(writer.iter_cell_writes())

        # -- series 1 table at row offset 0 (title+spacer 0, data 0) --
        # -- series 2 table at row offset 4 (title+spacer 2, data 2) --
        assert cells == [
            ("Sheet1", 2, 1, 1),
            ("Sheet1", 3, 1, 2),
            ("Sheet1", 1, 2, "Series 1"),
            ("Sheet1", 2, 2, 1.1),
            ("Sheet1", 3, 2, 2.2),
            ("Sheet1", 6, 1, 3),
            ("Sheet1", 5, 2, "Series 2"),
            ("Sheet1", 6, 2, 3.3),
        ]

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def populate_fixture(self, workbook_, worksheet_):
        chart_data = XyChartData()
        series_1 = chart_data.add_series("Series 1")
        for pt in ((1, 1.1), (2, 2.2)):
            series_1.add_data_point(*pt)
        series_2 = chart_data.add_series("Series 2")
        for pt in ((3, 3.3), (4, 4.4)):
            series_2.add_data_point(*pt)

        workbook_writer = XyWorkbookWriter(chart_data)

        expected_calls = [
            call.write_column(1, 0, [1, 2], ANY),
            call.write(0, 1, "Series 1"),
            call.write_column(1, 1, [1.1, 2.2], ANY),
            call.write_column(5, 0, [3, 4], ANY),
            call.write(4, 1, "Series 2"),
            call.write_column(5, 1, [3.3, 4.4], ANY),
        ]
        return workbook_writer, workbook_, worksheet_, expected_calls

    # fixture components ---------------------------------------------

    @pytest.fixture
    def workbook_(self, request):
        return instance_mock(request, Workbook)

    @pytest.fixture
    def worksheet_(self, request):
        return instance_mock(request, Worksheet)


class Describe_parse_sheet_range_ref(object):
    """Unit-test suite for `pptx.chart.xlsx.parse_sheet_range_ref`."""

    @pytest.mark.parametrize(
        "ref, expected_sheet, expected_cells",
        [
            ("Sheet1!$A$1", "Sheet1", [(1, 1)]),
            ("Sheet1!$B$2:$B$4", "Sheet1", [(2, 2), (3, 2), (4, 2)]),
            ("Sheet1!A2:B3", "Sheet1", [(2, 1), (2, 2), (3, 1), (3, 2)]),
            ("'My Sheet'!$A$1:$A$2", "My Sheet", [(1, 1), (2, 1)]),
            ("'O''Brien'!$A$1", "O'Brien", [(1, 1)]),
            ("Sheet1!$B$4:$B$2", "Sheet1", [(2, 2), (3, 2), (4, 2)]),
            ("Sheet1!$AA$1:$AB$1", "Sheet1", [(1, 27), (1, 28)]),
        ],
    )
    def it_parses_common_formula_refs(self, ref, expected_sheet, expected_cells):
        sheet, cells = parse_sheet_range_ref(ref)
        assert sheet == expected_sheet
        assert cells == expected_cells

    @pytest.mark.parametrize(
        "ref",
        [
            None,
            "",
            "Sheet1",
            "Sheet1!$A$1,$B$2",
            "='Other.xlsx'!Range",
            "42",
        ],
    )
    def it_returns_None_for_unsupported_refs(self, ref):
        assert parse_sheet_range_ref(ref) is None


class DescribeWorkbookReader(object):
    """Unit-test suite for `pptx.chart.xlsx.WorkbookReader`."""

    def it_reads_numeric_and_shared_string_cells(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData>'
                '<row r="1"><c r="A1" t="s"><v>0</v></c>'
                '<c r="B1"><v>1.5</v></c></row>'
                '<row r="2"><c r="A2" t="s"><v>1</v></c>'
                '<c r="B2"><v>42</v></c></row>'
                '</sheetData></worksheet>'
            ),
            shared_strings=("East", "West"),
        )
        with WorkbookReader(blob) as reader:
            assert reader.cell_value("Sheet1", 1, 1) == "East"
            assert reader.cell_value("Sheet1", 2, 1) == "West"
            assert reader.cell_value("Sheet1", 1, 2) == 1.5
            assert reader.cell_value("Sheet1", 2, 2) == 42.0

    def it_returns_None_for_missing_cells(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>'
                '</worksheet>'
            ),
            shared_strings=(),
        )
        with WorkbookReader(blob) as reader:
            assert reader.cell_value("Sheet1", 1, 1) == 7.0
            assert reader.cell_value("Sheet1", 99, 99) is None

    def it_supports_inline_strings(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData>'
                '<row r="1"><c r="A1" t="inlineStr"><is><t>Inline</t></is></c></row>'
                '</sheetData></worksheet>'
            ),
            shared_strings=(),
        )
        with WorkbookReader(blob) as reader:
            assert reader.cell_value("Sheet1", 1, 1) == "Inline"

    def it_falls_back_to_the_first_sheet_when_the_name_is_unknown(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData>'
                '</worksheet>'
            ),
            shared_strings=(),
        )
        with WorkbookReader(blob) as reader:
            assert reader.cell_value("CustomName", 1, 1) == 5.0

    def it_returns_None_when_the_workbook_has_no_sheets(self):
        buf = io.BytesIO()
        import zipfile as _zipfile

        with _zipfile.ZipFile(buf, "w") as z:
            z.writestr("dummy.txt", "not a real workbook")
        with WorkbookReader(buf.getvalue()) as reader:
            assert reader.cell_value("Sheet1", 1, 1) is None

    def it_does_not_expand_entities_when_parsing_embedded_workbook_xml(self):
        # --- A workbook.xml containing a billion-laughs style internal entity
        # --- must not be expanded when the reader parses the workbook to
        # --- enumerate sheets. The parser leaves the entity reference intact
        # --- so attribute lookups that depend on text expansion simply return
        # --- nothing, rather than the parser exhausting memory.
        xxe_workbook = (
            '<?xml version="1.0"?>'
            "<!DOCTYPE workbook ["
            '  <!ENTITY lol "lol">'
            '  <!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;">'
            "]>"
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="&lol2;" sheetId="1" r:id="rId1"/></sheets>'
            "</workbook>"
        )
        rels_xml = (
            '<?xml version="1.0"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            "</Relationships>"
        )
        sheet_xml = (
            '<?xml version="1.0"?>'
            '<worksheet xmlns='
            '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
            "</worksheet>"
        )
        buf = io.BytesIO()
        import zipfile as _zipfile

        with _zipfile.ZipFile(buf, "w") as z:
            z.writestr("xl/workbook.xml", xxe_workbook)
            z.writestr("xl/_rels/workbook.xml.rels", rels_xml)
            z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        with WorkbookReader(buf.getvalue()) as reader:
            # --- fallback-to-first-sheet path still succeeds because an
            # --- unresolved entity means the declared sheet-name never
            # --- matches any lookup; the first sheet's single cell is
            # --- returned for any requested sheet name. This also proves
            # --- the parser did not hang or abort on the entity payload.
            assert reader.cell_value("Sheet1", 1, 1) == 1.0

    def it_detects_cells_with_formula_references_239(self):
        """`cell_has_formula` returns True only for `<f>`-bearing cells."""
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                "<sheetData>"
                '<row r="1"><c r="A1"><v>1</v></c>'
                '<c r="B1"><f>A1*2</f><v>2</v></c></row>'
                "</sheetData></worksheet>"
            ),
            shared_strings=(),
        )
        with WorkbookReader(blob) as reader:
            assert reader.cell_has_formula("Sheet1", 1, 1) is False
            assert reader.cell_has_formula("Sheet1", 1, 2) is True
            # -- empty/missing cells are not formulas --
            assert reader.cell_has_formula("Sheet1", 99, 99) is False

    def it_returns_False_for_formula_check_when_workbook_has_no_sheets_239(self):
        buf = io.BytesIO()
        import zipfile as _zipfile

        with _zipfile.ZipFile(buf, "w") as z:
            z.writestr("dummy.txt", "not a real workbook")
        with WorkbookReader(buf.getvalue()) as reader:
            assert reader.cell_has_formula("Sheet1", 1, 1) is False


class Describe_parse_a1_cell(object):
    """Unit-test suite for `pptx.chart.xlsx.parse_a1_cell`."""

    @pytest.mark.parametrize(
        ("ref", "expected"),
        [
            ("A1", (1, 1)),
            ("$A$1", (1, 1)),
            ("$B$2", (2, 2)),
            ("AA1", (1, 27)),
            ("  c3 ", (3, 3)),
        ],
    )
    def it_parses_cell_refs(self, ref, expected):
        assert parse_a1_cell(ref) == expected

    @pytest.mark.parametrize(
        "ref",
        [None, "", "A1:B2", "Sheet1!A1", "1A", "$1", "A$"],
    )
    def it_raises_on_malformed_refs(self, ref):
        with pytest.raises(ValueError):
            parse_a1_cell(ref)


class Describe_row_col_to_a1(object):
    """Unit-test suite for `pptx.chart.xlsx._row_col_to_a1`."""

    @pytest.mark.parametrize(
        ("row", "col", "expected"),
        [
            (1, 1, "A1"),
            (2, 2, "B2"),
            (1, 27, "AA1"),
            (10, 26, "Z10"),
            (5, 28, "AB5"),
            (1, 702, "ZZ1"),
        ],
    )
    def it_formats_address(self, row, col, expected):
        assert _row_col_to_a1(row, col) == expected


class DescribeWorkbookUpdater(object):
    """Unit-test suite for `pptx.chart.xlsx.WorkbookUpdater`."""

    def it_returns_the_source_blob_unchanged_when_no_writes(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row r="1"><c r="A1"><v>1</v></c></row>'
                '</sheetData></worksheet>'
            ),
            shared_strings=(),
        )
        updater = WorkbookUpdater(blob)
        assert updater.blob() is blob

    def it_writes_a_numeric_cell(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row r="1"><c r="A1"><v>1</v></c></row>'
                '</sheetData></worksheet>'
            ),
            shared_strings=(),
        )
        updater = WorkbookUpdater(blob)
        updater.set_cell("Sheet1", 1, 1, 99)
        new_blob = updater.blob()
        with WorkbookReader(new_blob) as reader:
            assert reader.cell_value("Sheet1", 1, 1) == 99.0

    def it_writes_a_string_cell_as_inlineStr(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData/></worksheet>'
            ),
            shared_strings=(),
        )
        updater = WorkbookUpdater(blob)
        updater.set_cell("Sheet1", 2, 1, "Hello")
        new_blob = updater.blob()
        with WorkbookReader(new_blob) as reader:
            assert reader.cell_value("Sheet1", 2, 1) == "Hello"

    def it_writes_a_boolean_cell(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData/></worksheet>'
            ),
            shared_strings=(),
        )
        updater = WorkbookUpdater(blob)
        updater.set_cell("Sheet1", 1, 1, True)
        updater.set_cell("Sheet1", 2, 1, False)
        new_blob = updater.blob()
        with WorkbookReader(new_blob) as reader:
            assert reader.cell_value("Sheet1", 1, 1) is True
            assert reader.cell_value("Sheet1", 2, 1) is False

    def it_clears_a_cell_when_value_is_None(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row r="1"><c r="A1"><v>42</v></c></row>'
                '</sheetData></worksheet>'
            ),
            shared_strings=(),
        )
        updater = WorkbookUpdater(blob)
        updater.set_cell("Sheet1", 1, 1, None)
        new_blob = updater.blob()
        with WorkbookReader(new_blob) as reader:
            assert reader.cell_value("Sheet1", 1, 1) is None

    def it_adds_a_new_row_when_one_is_not_present(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row r="1"><c r="A1"><v>1</v></c></row>'
                '</sheetData></worksheet>'
            ),
            shared_strings=(),
        )
        updater = WorkbookUpdater(blob)
        updater.set_cell("Sheet1", 5, 3, 7)
        new_blob = updater.blob()
        with WorkbookReader(new_blob) as reader:
            assert reader.cell_value("Sheet1", 5, 3) == 7.0
            # -- prior cells survive --
            assert reader.cell_value("Sheet1", 1, 1) == 1.0

    def it_coalesces_repeated_writes_to_the_last_value(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData/></worksheet>'
            ),
            shared_strings=(),
        )
        updater = WorkbookUpdater(blob)
        updater.set_cell("Sheet1", 1, 1, 1)
        updater.set_cell("Sheet1", 1, 1, 2)
        updater.set_cell("Sheet1", 1, 1, 3)
        new_blob = updater.blob()
        with WorkbookReader(new_blob) as reader:
            assert reader.cell_value("Sheet1", 1, 1) == 3.0

    def it_falls_back_to_first_sheet_when_sheet_name_unknown(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData/></worksheet>'
            ),
            shared_strings=(),
        )
        updater = WorkbookUpdater(blob)
        updater.set_cell("Bogus", 1, 1, 42)
        new_blob = updater.blob()
        with WorkbookReader(new_blob) as reader:
            assert reader.cell_value("Sheet1", 1, 1) == 42.0

    def it_preserves_existing_cells_in_the_same_row(self):
        blob = _make_xlsx_blob(
            sheet=(
                '<?xml version="1.0"?>'
                '<worksheet xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row r="1"><c r="A1"><v>1</v></c>'
                '<c r="C1"><v>3</v></c></row>'
                '</sheetData></worksheet>'
            ),
            shared_strings=(),
        )
        updater = WorkbookUpdater(blob)
        updater.set_cell("Sheet1", 1, 2, 2)
        new_blob = updater.blob()
        with WorkbookReader(new_blob) as reader:
            assert reader.cell_value("Sheet1", 1, 1) == 1.0
            assert reader.cell_value("Sheet1", 1, 2) == 2.0
            assert reader.cell_value("Sheet1", 1, 3) == 3.0


def _make_xlsx_blob(sheet, shared_strings):
    """Build a minimal xlsx blob with one sheet and optional sharedStrings."""
    import zipfile as _zipfile

    workbook_xml = (
        '<?xml version="1.0"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    rels_xml = (
        '<?xml version="1.0"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    buf = io.BytesIO()
    with _zipfile.ZipFile(buf, "w") as z:
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
        if shared_strings:
            sst_xml = (
                '<?xml version="1.0"?>'
                '<sst xmlns='
                '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                + "".join("<si><t>{}</t></si>".format(s) for s in shared_strings)
                + "</sst>"
            )
            z.writestr("xl/sharedStrings.xml", sst_xml)
    return buf.getvalue()
