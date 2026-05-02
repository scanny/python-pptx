"""Chart builder and related objects."""

from __future__ import annotations

import io
import re
import zipfile
from contextlib import contextmanager

from lxml import etree
from xlsxwriter import Workbook


class _BaseWorkbookWriter(object):
    """Base class for workbook writers, providing shared members."""

    def __init__(self, chart_data):
        super(_BaseWorkbookWriter, self).__init__()
        self._chart_data = chart_data

    @property
    def xlsx_blob(self):
        """bytes for Excel file containing chart_data."""
        xlsx_file = io.BytesIO()
        with self._open_worksheet(xlsx_file) as (workbook, worksheet):
            self._populate_worksheet(workbook, worksheet)
        return xlsx_file.getvalue()

    @contextmanager
    def _open_worksheet(self, xlsx_file):
        """
        Enable XlsxWriter Worksheet object to be opened, operated on, and
        then automatically closed within a `with` statement. A filename or
        stream object (such as an `io.BytesIO` instance) is expected as
        *xlsx_file*.
        """
        workbook = Workbook(xlsx_file, {"in_memory": True})
        worksheet = workbook.add_worksheet()
        yield workbook, worksheet
        workbook.close()

    def _populate_worksheet(self, workbook, worksheet):
        """
        Must be overridden by each subclass to provide the particulars of
        writing the spreadsheet data.
        """
        raise NotImplementedError("must be provided by each subclass")


class CategoryWorkbookWriter(_BaseWorkbookWriter):
    """
    Determines Excel worksheet layout and can write an Excel workbook from
    a CategoryChartData object. Serves as the authority for Excel worksheet
    ranges.
    """

    @property
    def categories_ref(self):
        """
        The Excel worksheet reference to the categories for this chart (not
        including the column heading).
        """
        categories = self._chart_data.categories
        if categories.depth == 0:
            raise ValueError("chart data contains no categories")
        right_col = chr(ord("A") + categories.depth - 1)
        bottom_row = categories.leaf_count + 1
        return "Sheet1!$A$2:$%s$%d" % (right_col, bottom_row)

    def series_name_ref(self, series):
        """
        Return the Excel worksheet reference to the cell containing the name
        for *series*. This also serves as the column heading for the series
        values.
        """
        return "Sheet1!$%s$1" % self._series_col_letter(series)

    def values_ref(self, series):
        """
        The Excel worksheet reference to the values for this series (not
        including the column heading).
        """
        return "Sheet1!${col_letter}$2:${col_letter}${bottom_row}".format(
            **{
                "col_letter": self._series_col_letter(series),
                "bottom_row": len(series) + 1,
            }
        )

    @staticmethod
    def _column_reference(column_number):
        """Return str Excel column reference like 'BQ' for *column_number*.

        *column_number* is an int in the range 1-16384 inclusive, where
        1 maps to column 'A'.
        """
        if column_number < 1 or column_number > 16384:
            raise ValueError("column_number must be in range 1-16384")

        # ---Work right-to-left, one order of magnitude at a time. Note there
        #    is no zero representation in Excel address scheme, so this is
        #    not just a conversion to base-26---

        col_ref = ""
        while column_number:
            remainder = column_number % 26
            if remainder == 0:
                remainder = 26

            col_letter = chr(ord("A") + remainder - 1)
            col_ref = col_letter + col_ref

            # ---Advance to next order of magnitude or terminate loop. The
            # minus-one in this expression reflects the fact the next lower
            # order of magnitude has a minumum value of 1 (not zero). This is
            # essentially the complement to the "if it's 0 make it 26' step
            # above.---
            column_number = (column_number - 1) // 26

        return col_ref

    def _populate_worksheet(self, workbook, worksheet):
        """
        Write the chart data contents to *worksheet* in category chart
        layout. Write categories starting in the first column starting in
        the second row, and proceeding one column per category level (for
        charts having multi-level categories). Write series as columns
        starting in the next following column, placing the series title in
        the first cell.
        """
        self._write_categories(workbook, worksheet)
        self._write_series(workbook, worksheet)

    def _series_col_letter(self, series):
        """
        The letter of the Excel worksheet column in which the data for a
        series appears.
        """
        column_number = 1 + series.categories.depth + series.index
        return self._column_reference(column_number)

    def _write_categories(self, workbook, worksheet):
        """
        Write the categories column(s) to *worksheet*. Categories start in
        the first column starting in the second row, and proceeding one
        column per category level (for charts having multi-level categories).
        A date category is formatted as a date. All others are formatted
        `General`.
        """
        categories = self._chart_data.categories
        num_format = workbook.add_format({"num_format": categories.number_format})
        depth = categories.depth
        for idx, level in enumerate(categories.levels):
            col = depth - idx - 1
            self._write_cat_column(worksheet, col, level, num_format)

    def _write_cat_column(self, worksheet, col, level, num_format):
        """
        Write a category column defined by *level* to *worksheet* at offset
        *col* and formatted with *num_format*.
        """
        worksheet.set_column(col, col, 10)  # wide enough for a date
        for off, name in level:
            row = off + 1
            worksheet.write(row, col, name, num_format)

    def _write_series(self, workbook, worksheet):
        """
        Write the series column(s) to *worksheet*. Series start in the column
        following the last categories column, placing the series title in the
        first cell.
        """
        col_offset = self._chart_data.categories.depth
        for idx, series in enumerate(self._chart_data):
            num_format = workbook.add_format({"num_format": series.number_format})
            series_col = idx + col_offset
            worksheet.write(0, series_col, series.name)
            worksheet.write_column(1, series_col, series.values, num_format)


class XyWorkbookWriter(_BaseWorkbookWriter):
    """
    Determines Excel worksheet layout and can write an Excel workbook from XY
    chart data. Serves as the authority for Excel worksheet ranges.
    """

    def series_name_ref(self, series):
        """
        Return the Excel worksheet reference to the cell containing the name
        for *series*. This also serves as the column heading for the series
        Y values.
        """
        row = self.series_table_row_offset(series) + 1
        return "Sheet1!$B$%d" % row

    def series_table_row_offset(self, series):
        """
        Return the number of rows preceding the data table for *series* in
        the Excel worksheet.
        """
        title_and_spacer_rows = series.index * 2
        data_point_rows = series.data_point_offset
        return title_and_spacer_rows + data_point_rows

    def x_values_ref(self, series):
        """
        The Excel worksheet reference to the X values for this chart (not
        including the column label).
        """
        top_row = self.series_table_row_offset(series) + 2
        bottom_row = top_row + len(series) - 1
        return "Sheet1!$A$%d:$A$%d" % (top_row, bottom_row)

    def y_values_ref(self, series):
        """
        The Excel worksheet reference to the Y values for this chart (not
        including the column label).
        """
        top_row = self.series_table_row_offset(series) + 2
        bottom_row = top_row + len(series) - 1
        return "Sheet1!$B$%d:$B$%d" % (top_row, bottom_row)

    def _populate_worksheet(self, workbook, worksheet):
        """
        Write chart data contents to *worksheet* in the standard XY chart
        layout. Write the data for each series to a separate two-column
        table, X values in column A and Y values in column B. Place the
        series label in the first (heading) cell of the column.
        """
        chart_num_format = workbook.add_format({"num_format": self._chart_data.number_format})
        for series in self._chart_data:
            series_num_format = workbook.add_format({"num_format": series.number_format})
            offset = self.series_table_row_offset(series)
            # write X values
            worksheet.write_column(offset + 1, 0, series.x_values, chart_num_format)
            # write Y values
            worksheet.write(offset, 1, series.name)
            worksheet.write_column(offset + 1, 1, series.y_values, series_num_format)


class BubbleWorkbookWriter(XyWorkbookWriter):
    """
    Service object that knows how to write an Excel workbook from bubble
    chart data.
    """

    def bubble_sizes_ref(self, series):
        """
        The Excel worksheet reference to the range containing the bubble
        sizes for *series* (not including the column heading cell).
        """
        top_row = self.series_table_row_offset(series) + 2
        bottom_row = top_row + len(series) - 1
        return "Sheet1!$C$%d:$C$%d" % (top_row, bottom_row)

    def _populate_worksheet(self, workbook, worksheet):
        """
        Write chart data contents to *worksheet* in the bubble chart layout.
        Write the data for each series to a separate three-column table with
        X values in column A, Y values in column B, and bubble sizes in
        column C. Place the series label in the first (heading) cell of the
        values column.
        """
        chart_num_format = workbook.add_format({"num_format": self._chart_data.number_format})
        for series in self._chart_data:
            series_num_format = workbook.add_format({"num_format": series.number_format})
            offset = self.series_table_row_offset(series)
            # write X values
            worksheet.write_column(offset + 1, 0, series.x_values, chart_num_format)
            # write Y values
            worksheet.write(offset, 1, series.name)
            worksheet.write_column(offset + 1, 1, series.y_values, series_num_format)
            # write bubble sizes
            worksheet.write(offset, 2, "Size")
            worksheet.write_column(offset + 1, 2, series.bubble_sizes, chart_num_format)


# --- Excel range reference parsing --------------------------------------------------

# e.g. "Sheet1!$A$2:$A$5", "Sheet1!$B$1", "'My Sheet'!$A$1:$B$3"
_sheet_ref_re = re.compile(
    r"""
    ^\s*
    (?:'(?P<qsheet>(?:[^']|'')*)'|(?P<sheet>[^!]+))
    !
    \$?(?P<col1>[A-Za-z]+)\$?(?P<row1>\d+)
    (?:
        :\$?(?P<col2>[A-Za-z]+)\$?(?P<row2>\d+)
    )?
    \s*$
    """,
    re.VERBOSE,
)


def _column_letters_to_index(letters):
    """Return 1-based column index for Excel column letters like 'A', 'Z', 'AA'."""
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def parse_sheet_range_ref(ref):
    """Return `(sheet_name, [(row, col), ...])` for A1-style `ref`.

    Rows and columns are 1-based. The returned list enumerates cells in
    row-major order (top-to-bottom, then left-to-right within the range). For
    a single-cell reference the list has exactly one element. Returns |None|
    for `ref` values that don't match the expected pattern (e.g. multi-range
    references joined with commas or broken inputs) so that callers can skip
    rewriting caches they can't resolve.
    """
    if ref is None:
        return None
    m = _sheet_ref_re.match(ref)
    if m is None:
        return None
    sheet = m.group("qsheet")
    if sheet is not None:
        sheet = sheet.replace("''", "'")
    else:
        sheet = m.group("sheet")
    col1 = _column_letters_to_index(m.group("col1"))
    row1 = int(m.group("row1"))
    col2 = _column_letters_to_index(m.group("col2")) if m.group("col2") else col1
    row2 = int(m.group("row2")) if m.group("row2") else row1
    if row2 < row1:
        row1, row2 = row2, row1
    if col2 < col1:
        col1, col2 = col2, col1
    cells = [(r, c) for r in range(row1, row2 + 1) for c in range(col1, col2 + 1)]
    return sheet, cells


# --- Workbook reader ---------------------------------------------------------------

_SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_OPC_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_OFC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class WorkbookReader(object):
    """Minimal read-only view of an embedded chart workbook.

    Parses an `.xlsx` blob using `zipfile` + `lxml` and exposes a
    `cell_value(sheet_name, row, col)` lookup that returns the cell value as
    a `str`, `float`, or |None| when the cell is empty. Only the first
    worksheet of the workbook is required for the chart-cache refresh
    use-case; additional worksheets are parsed lazily on access.

    This is intentionally minimal: it supports inline strings, shared
    strings, and raw numeric values — the forms that PowerPoint and
    `XlsxWriter` actually emit for chart data. Formula results are read from
    the cached `<c:v>` child when present (XlsxWriter does not emit
    formulas).
    """

    def __init__(self, xlsx_blob):
        self._xlsx_blob = xlsx_blob
        self._shared_strings = None
        self._sheet_names = None  # list[str] ordered by workbook.xml
        self._sheet_targets = None  # dict[str, str] sheet_name -> internal path
        self._sheet_cache = {}  # dict[str, dict[(row,col), value]]
        self._zf = None

    def __enter__(self):
        self._open()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()

    def _open(self):
        if self._zf is None:
            self._zf = zipfile.ZipFile(io.BytesIO(self._xlsx_blob))

    def close(self):
        if self._zf is not None:
            self._zf.close()
            self._zf = None

    # -- public API ---------------------------------------------------------

    def cell_value(self, sheet_name, row, col):
        """Return the value of cell (`row`, `col`) in `sheet_name`, or |None|.

        Rows and columns are 1-based. Returns |None| when the sheet is not
        present in the workbook, when the cell has no value, or when the
        cell's stored value is an empty string. Numeric cells are returned
        as `float`, string cells (inline or shared) as `str`.
        """
        self._open()
        self._load_workbook_meta()
        sheet_cells = self._load_sheet(sheet_name)
        if sheet_cells is None:
            return None
        return sheet_cells.get((row, col))

    # -- internal helpers --------------------------------------------------

    def _load_workbook_meta(self):
        if self._sheet_names is not None:
            return
        # --- read workbook.xml for sheet order + rIds ---
        try:
            wb_bytes = self._zf.read("xl/workbook.xml")
        except KeyError:
            self._sheet_names = []
            self._sheet_targets = {}
            return
        wb = etree.fromstring(wb_bytes)
        sheet_elms = wb.findall(f"{{{_SML_NS}}}sheets/{{{_SML_NS}}}sheet")
        names = []
        rid_by_name = {}
        for s in sheet_elms:
            name = s.get("name")
            rid = s.get(f"{{{_OFC_REL_NS}}}id")
            names.append(name)
            rid_by_name[name] = rid
        # --- resolve rIds via xl/_rels/workbook.xml.rels ---
        try:
            rels_bytes = self._zf.read("xl/_rels/workbook.xml.rels")
        except KeyError:
            rels_bytes = b""
        target_by_rid = {}
        if rels_bytes:
            rels = etree.fromstring(rels_bytes)
            for rel in rels.findall(f"{{{_OPC_REL_NS}}}Relationship"):
                target_by_rid[rel.get("Id")] = rel.get("Target")
        self._sheet_names = names
        self._sheet_targets = {}
        for name, rid in rid_by_name.items():
            target = target_by_rid.get(rid)
            if target is None:
                continue
            # Targets are relative to xl/ (e.g. 'worksheets/sheet1.xml').
            if target.startswith("/"):
                path = target.lstrip("/")
            else:
                path = "xl/" + target
            self._sheet_targets[name] = path

    def _load_shared_strings(self):
        if self._shared_strings is not None:
            return
        try:
            ss_bytes = self._zf.read("xl/sharedStrings.xml")
        except KeyError:
            self._shared_strings = []
            return
        sst = etree.fromstring(ss_bytes)
        results = []
        for si in sst.findall(f"{{{_SML_NS}}}si"):
            # --- si may have a single <t> or a sequence of <r><t>… runs ---
            parts = []
            for t in si.iter(f"{{{_SML_NS}}}t"):
                parts.append(t.text or "")
            results.append("".join(parts))
        self._shared_strings = results

    def _load_sheet(self, sheet_name):
        if sheet_name in self._sheet_cache:
            return self._sheet_cache[sheet_name]
        self._load_workbook_meta()
        # --- fall back to first sheet if named sheet not present ---
        # (PowerPoint/XlsxWriter always put chart data in Sheet1; real-world
        # files sometimes keep the literal "Sheet1" reference even when the
        # tab has been renamed.) ---
        path = self._sheet_targets.get(sheet_name)
        if path is None and self._sheet_names:
            path = self._sheet_targets.get(self._sheet_names[0])
        if path is None:
            self._sheet_cache[sheet_name] = None
            return None
        try:
            sheet_bytes = self._zf.read(path)
        except KeyError:
            self._sheet_cache[sheet_name] = None
            return None
        self._load_shared_strings()
        cells = {}
        ws = etree.fromstring(sheet_bytes)
        for c in ws.iter(f"{{{_SML_NS}}}c"):
            addr = c.get("r")
            if not addr:
                continue
            m = re.match(r"([A-Za-z]+)(\d+)$", addr)
            if m is None:
                continue
            col = _column_letters_to_index(m.group(1))
            row = int(m.group(2))
            t = c.get("t")
            v = c.find(f"{{{_SML_NS}}}v")
            if t == "inlineStr":
                is_elm = c.find(f"{{{_SML_NS}}}is")
                text = ""
                if is_elm is not None:
                    for t_elm in is_elm.iter(f"{{{_SML_NS}}}t"):
                        text += t_elm.text or ""
                cells[(row, col)] = text
                continue
            if v is None or v.text is None or v.text == "":
                continue
            raw = v.text
            if t == "s":
                try:
                    idx = int(raw)
                    cells[(row, col)] = self._shared_strings[idx]
                except (ValueError, IndexError):
                    cells[(row, col)] = raw
            elif t == "str":
                cells[(row, col)] = raw
            elif t == "b":
                cells[(row, col)] = raw != "0"
            elif t == "e":
                # --- error cells: return the error text so callers can
                #     surface it verbatim. ---
                cells[(row, col)] = raw
            else:
                # --- default numeric ---
                try:
                    cells[(row, col)] = float(raw)
                except ValueError:
                    cells[(row, col)] = raw
        self._sheet_cache[sheet_name] = cells
        return cells
