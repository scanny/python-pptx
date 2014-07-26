# encoding: utf-8

"""
Series-related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..dml.fill import FillFormat
from ..dml.line import LineFormat
from ..oxml.ns import qn
from ..util import lazyproperty


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
    @lazyproperty
    def fill(self):
        """
        |FillFormat| instance for this series, providing access to fill
        properties such as fill color.
        """
        spPr = self._element.get_or_add_spPr()
        return FillFormat.from_fill_parent(spPr)

    def get_or_add_ln(self):
        """
        Return the ``<a:ln>`` element containing the line format properties
        XML for this shape. Part of the callback interface required by
        |LineFormat|.
        """
        spPr = self._element.get_or_add_spPr()
        ln = spPr.get_or_add_ln()
        return ln

    @property
    def invert_if_negative(self):
        """
        |True| if the fill of a bar having a negative value should be
        inverted. |False| if the fill should be the same regardless of the
        bar's value. When |True|, a bar with a solid fill appears with white
        fill; in a bar with gradient fill, the direction of the gradient is
        reversed, e.g. dark -> light instead of light -> dark.
        """
        invertIfNegative = self._element.invertIfNegative
        if invertIfNegative is None:
            return True
        return invertIfNegative.val

    @invert_if_negative.setter
    def invert_if_negative(self, value):
        invertIfNegative = self._element.get_or_add_invertIfNegative()
        invertIfNegative.val = value

    @lazyproperty
    def line(self):
        """
        |LineFormat| instance for this shape, providing access to line
        properties such as line color.
        """
        return LineFormat(self)

    @property
    def ln(self):
        """
        The ``<a:ln>`` element containing the line format properties such as
        line color and width. |None| if no ``<a:ln>`` element is present.
        Part of the callback interface required by |LineFormat|.
        """
        spPr = self._element.spPr
        if spPr is None:
            return None
        return spPr.ln


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
