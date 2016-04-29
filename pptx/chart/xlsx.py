# encoding: utf-8

"""
Chart builder and related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from contextlib import contextmanager

from xlsxwriter import Workbook

from ..compat import BytesIO


class WorkbookWriter(object):
    """
    Service object that knows how to write an Excel workbook for chart data.
    """
    @classmethod
    def xlsx_blob(cls, categories, series):
        """
        Return the byte stream of an Excel file formatted as chart data for
        a chart having *categories* and *series*.
        """
        xlsx_file = BytesIO()
        with cls._open_worksheet(xlsx_file) as (workbook, worksheet):
            cls._populate_worksheet(workbook, worksheet, categories, series)
        return xlsx_file.getvalue()

    @staticmethod
    @contextmanager
    def _open_worksheet(xlsx_file):
        """
        Enable XlsxWriter Worksheet object to be opened, operated on, and
        then automatically closed within a `with` statement. A filename or
        stream object (such as a ``BytesIO`` instance) is expected as
        *xlsx_file*.
        """
        workbook = Workbook(xlsx_file, {'in_memory': True})
        worksheet = workbook.add_worksheet()
        yield workbook, worksheet
        workbook.close()

    @classmethod
    def _populate_worksheet(cls, workbook, worksheet, categories, series):
        """
        Write *categories* and *series* to *worksheet* in the standard
        layout, categories in first column starting in second row, and series
        as columns starting in second column, series title in first cell.
        Make the whole range an Excel List.
        """
        worksheet.write_column(1, 0, categories)
        for series in series:
            num_format = (
                workbook.add_format({'num_format': series.number_format})
            )
            series_col = series.index + 1
            worksheet.write(0, series_col, series.name)
            worksheet.write_column(1, series_col, series.values, num_format)


class XyWorkbookWriter(object):
    """
    Determines Excel worksheet layout and can write an Excel workbook from XY
    chart data. Serves as the authority for Excel worksheet ranges.
    """
    def __init__(self, chart_data):
        super(XyWorkbookWriter, self).__init__()
        self._chart_data = chart_data

    @contextmanager
    def _open_worksheet(self, xlsx_file):
        """
        Enable XlsxWriter Worksheet object to be opened, operated on, and
        then automatically closed within a `with` statement. A filename or
        stream object (such as a ``BytesIO`` instance) is expected as
        *xlsx_file*.
        """
        workbook = Workbook(xlsx_file, {'in_memory': True})
        worksheet = workbook.add_worksheet()
        yield workbook, worksheet
        workbook.close()

    def series_name_ref(self, series):
        """
        Return the Excel worksheet reference to the cell containing the name
        for *series*. This also serves as the column heading for the series
        Y values.
        """
        row = self.series_table_row_offset(series) + 1
        return 'Sheet1!$B$%d' % row

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

    @property
    def xlsx_blob(self):
        """
        Return the byte stream of an Excel file formatted as chart data for
        the XY chart specified in *chart_data*.
        """
        xlsx_file = BytesIO()
        with self._open_worksheet(xlsx_file) as (workbook, worksheet):
            self._populate_worksheet(workbook, worksheet)
        return xlsx_file.getvalue()

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
        for series in self._chart_data:
            offset = self.series_table_row_offset(series)
            # write X values
            worksheet.write_column(offset+1, 0, series.x_values)
            # write Y values
            worksheet.write(offset, 1, series.name)
            worksheet.write_column(offset+1, 1, series.y_values)


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
        raise NotImplementedError
