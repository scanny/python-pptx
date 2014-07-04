# encoding: utf-8

"""
Chart shape-related objects such as Chart.
"""

from __future__ import absolute_import, print_function, unicode_literals


class Chart(object):
    """
    A chart object.
    """
    def __init__(self, chartSpace):
        super(Chart, self).__init__()
        self._chartSpace = chartSpace
