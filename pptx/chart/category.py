# encoding: utf-8

"""
Category-related objects. The |Categories| object is returned by
``Plot.categories`` and contains zero or more |Category| objects, each
representing one of the category labels associated with the plot. Categories
can be hierarchical, so there are members allowing discovery of the depth of
that hierarchy and means to navigate it.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from collections import Sequence


class Categories(Sequence):
    """
    A sequence of |Category| objects, each representing a category label on
    the chart. Provides properties for dealing with hierarchical categories.
    """
    def __init__(self, xChart):
        super(Categories, self).__init__()
        self._xChart = xChart

    def __getitem__(self, idx):
        raise NotImplementedError

    def __len__(self):
        # a category can be "null", meaning the Excel cell for it is empty.
        # In this case, there is no c:pt element for it. The "empty" category
        # will, however, be accounted for in c:cat//c:ptCount/@val, which
        # reflects the true length of the categories collection.
        return self._xChart.cat_pt_count
