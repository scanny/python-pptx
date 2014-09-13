# encoding: utf-8

"""
Series-related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from collections import Sequence

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

    @property
    def index(self):
        """
        The zero-based integer index of this series as reported in its
        `c:ser/c:idx` element.
        """
        return self._element.idx.val

    @property
    def name(self):
        """
        The string label given to this series, appears as the title of the
        column for this series in the Excel worksheet. It also appears as the
        label for this series in the legend.
        """
        names = self._element.xpath('./c:tx//c:pt/c:v/text()')
        name = names[0] if names else ''
        return name

    @property
    def values(self):
        """
        Read-only. A sequence containing the float values for this series, in
        the order they appear on the chart.
        """
        ser = self._element
        value_pt_elms = ser.val_pts
        return tuple(pt.value for pt in value_pt_elms)


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
        |True| if a point having a value less than zero should appear with a
        fill different than those with a positive value. |False| if the fill
        should be the same regardless of the bar's value. When |True|, a bar
        with a solid fill appears with white fill; in a bar with gradient
        fill, the direction of the gradient is reversed, e.g. dark -> light
        instead of light -> dark. The term "invert" here should be understood
        to mean "invert the *direction* of the *fill gradient*".
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
        properties such as line color and width.
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
    @property
    def smooth(self):
        """
        Read/write boolean specifying whether to use curve smoothing to
        form the line connecting the data points in this series into
        a continuous curve. If |False|, a series of straight line segments
        are used to connect the points.
        """
        smooth = self._element.smooth
        if smooth is None:
            return True
        return smooth.val

    @smooth.setter
    def smooth(self, value):
        self._element.get_or_add_smooth().val = value


class PieSeries(_BaseSeries):
    """
    A data point series belonging to a pie plot.
    """


class SeriesCollection(Sequence):
    """
    A sequence of |Series| objects.
    """
    def __init__(self, parent_elm):
        # *parent_elm* can be either a c:chartSpace or xChart element
        super(SeriesCollection, self).__init__()
        self._element = parent_elm

    def __getitem__(self, index):
        ser = self._element.sers[index]
        return _SeriesFactory(ser)

    def __len__(self):
        return len(self._element.sers)


def _SeriesFactory(ser):
    """
    Return an instance of the appropriate subclass of _BaseSeries based on the
    xChart element *ser* appears in.
    """
    xChart_tag = ser.getparent().tag
    if xChart_tag == qn('c:barChart'):
        return BarSeries(ser)
    if xChart_tag == qn('c:lineChart'):
        return LineSeries(ser)
    if xChart_tag == qn('c:pieChart'):
        return PieSeries(ser)
    raise ValueError('unsupported series type %s' % xChart_tag)
