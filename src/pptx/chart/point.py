"""Data point-related objects."""

from __future__ import annotations

from collections.abc import Sequence

from pptx.chart.datalabel import DataLabel
from pptx.chart.marker import Marker
from pptx.dml.chtfmt import ChartFormat
from pptx.util import lazyproperty


class _BasePoints(Sequence):
    """
    Sequence providing access to the individual data points in a series.
    """

    def __init__(self, ser):
        super(_BasePoints, self).__init__()
        self._element = ser
        self._ser = ser

    def __getitem__(self, idx):
        if idx < 0 or idx >= self.__len__():
            raise IndexError("point index out of range")
        return Point(self._ser, idx)


class BubblePoints(_BasePoints):
    """
    Sequence providing access to the individual data points in
    a |BubbleSeries| object.
    """

    def __len__(self):
        return min(
            self._ser.xVal_ptCount_val,
            self._ser.yVal_ptCount_val,
            self._ser.bubbleSize_ptCount_val,
        )


class CategoryPoints(_BasePoints):
    """
    Sequence providing access to individual |Point| objects, each
    representing the visual properties of a data point in the specified
    category series.
    """

    def __len__(self):
        return self._ser.cat_ptCount_val


class Point(object):
    """
    Provides access to the properties of an individual data point in
    a series, such as the visual properties of its marker and the text and
    font of its data label.
    """

    def __init__(self, ser, idx):
        super(Point, self).__init__()
        self._element = ser
        self._ser = ser
        self._idx = idx

    @lazyproperty
    def data_label(self):
        """
        The |DataLabel| object representing the label on this data point.
        """
        return DataLabel(self._ser, self._idx)

    @lazyproperty
    def format(self):
        """
        The |ChartFormat| object providing access to the shape formatting
        properties of this data point, such as line and fill.
        """
        dPt = self._ser.get_or_add_dPt_for_point(self._idx)
        return ChartFormat(dPt)

    @property
    def invert_if_negative(self):
        """
        Read/write bool. |True| if this data point should appear with a
        different fill when it has a negative value. Overrides the series-
        level setting when present. When |True|, a data point with a solid
        fill appears with white fill; the direction of a gradient fill is
        reversed. Returns |True| when no explicit `c:invertIfNegative`
        element is present on this data point, which is the default
        behavior.

        .. note::

           When assigning a solid fill to a data point that represents a
           negative value, PowerPoint will render that bar as **white**
           unless ``invert_if_negative`` is explicitly set to |False| on
           that point. The inversion default is |True| at the point level,
           so the explicit fill is *inverted* to white by PowerPoint. To
           show your chosen fill color on a negative bar, pair the fill
           assignment with ``point.invert_if_negative = False``::

               point = plot.series[0].points[i]
               fill = point.format.fill
               fill.solid()
               fill.fore_color.rgb = RGBColor(0xF3, 0x5D, 0x5D)
               point.invert_if_negative = False  # preserve red on negatives
        """
        dPt = self._ser.get_or_add_dPt_for_point(self._idx)
        invertIfNegative = dPt.invertIfNegative
        if invertIfNegative is None:
            return True
        return invertIfNegative.val

    @invert_if_negative.setter
    def invert_if_negative(self, value):
        dPt = self._ser.get_or_add_dPt_for_point(self._idx)
        invertIfNegative = dPt.get_or_add_invertIfNegative()
        invertIfNegative.val = value

    @lazyproperty
    def marker(self):
        """
        The |Marker| instance for this point, providing access to the visual
        properties of the data point marker, such as fill and line. Setting
        these properties overrides any value set at the series level.
        """
        dPt = self._ser.get_or_add_dPt_for_point(self._idx)
        return Marker(dPt)


class XyPoints(_BasePoints):
    """
    Sequence providing access to the individual data points in an |XySeries|
    object.
    """

    def __len__(self):
        return min(self._ser.xVal_ptCount_val, self._ser.yVal_ptCount_val)
