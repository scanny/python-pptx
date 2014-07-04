# encoding: utf-8

"""
Axis-related chart objects.
"""

from __future__ import absolute_import, print_function, unicode_literals


class _BaseAxis(object):
    """
    Base class for chart axis classes.
    """
    def __init__(self, xAx_elm):
        super(_BaseAxis, self).__init__()
        self._element = xAx_elm


class CategoryAxis(_BaseAxis):
    """
    A category axis of a chart.
    """


class ValueAxis(_BaseAxis):
    """
    A value axis of a chart.
    """
