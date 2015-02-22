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
