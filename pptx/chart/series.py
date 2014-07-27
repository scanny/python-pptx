# encoding: utf-8

"""
Series-related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..oxml.ns import qn


class _BaseSeries(object):
    """
    Base class for |BarSeries| and other series classes.
    """
    def __init__(self, ser):
        super(_BaseSeries, self).__init__()
        self._element = ser


class BarSeries(_BaseSeries):
    """
    A data point series belonging to a bar plot.
    """


class LineSeries(_BaseSeries):
    """
    A data point series belonging to a line plot.
    """


class PieSeries(_BaseSeries):
    """
    A data point series belonging to a pie plot.
    """


def SeriesFactory(plot_elm, ser):
    """
    Return an instance of the appropriate subclass of _BaseSeries based on the
    tagname of *plot_elm*.
    """
    if plot_elm.tag == qn('c:barChart'):
        return BarSeries(ser)
    if plot_elm.tag == qn('c:lineChart'):
        return LineSeries(ser)
    if plot_elm.tag == qn('c:pieChart'):
        return PieSeries(ser)
    raise ValueError('unsupported series type %s' % plot_elm.tag)
