# encoding: utf-8

"""
Chart builder and related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals


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
        raise NotImplementedError
