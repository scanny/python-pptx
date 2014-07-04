# encoding: utf-8

"""
Chart shape-related objects such as Chart.
"""

from __future__ import absolute_import, print_function, unicode_literals

from .axis import CategoryAxis


class Chart(object):
    """
    A chart object.
    """
    def __init__(self, chartSpace, chart_part):
        super(Chart, self).__init__()
        self._chartSpace = chartSpace
        self._chart_part = chart_part

    @property
    def category_axis(self):
        """
        The category axis of this chart. Raises |ValueError| if no category
        axis is defined.
        """
        catAx = self._chartSpace.catAx
        if catAx is None:
            raise ValueError('chart has no category axis')
        return CategoryAxis(catAx)
