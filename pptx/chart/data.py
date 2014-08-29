# encoding: utf-8

"""
ChartData and related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals


class ChartData(object):
    """
    Primarily a data transfer object, `ChartData` accumulates data specifying
    the categories and series values for a plot. However, it also serves as
    a broker to the services of |ChartXmlWriter| and |WorkbookWriter| for
    obtaining chart XML and Excel objects respectively.
    """
    def __init__(self):
        super(ChartData, self).__init__()
        self._categories = []

    @property
    def categories(self):
        """
        An immutable sequence containing the category labels currently
        defined in this chart data object.
        """
        return tuple(self._categories)

    @categories.setter
    def categories(self, categories):
        # Contents need to be replaced in-place so reference sent to
        # _SeriesData objects retain access to latest values
        self._categories[:] = categories
