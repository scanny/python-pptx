# encoding: utf-8

"""
Axis-related chart objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..enum.chart import XL_TICK_MARK
from ..util import lazyproperty


class _BaseAxis(object):
    """
    Base class for chart axis classes.
    """
    def __init__(self, xAx_elm):
        super(_BaseAxis, self).__init__()
        self._element = xAx_elm

    @property
    def major_tick_mark(self):
        """
        Read/write :ref:`XlTickMark` value specifying the type of major tick
        mark for this axis.
        """
        majorTickMark = self._element.majorTickMark
        if majorTickMark is None:
            return XL_TICK_MARK.CROSS
        return majorTickMark.val

    @major_tick_mark.setter
    def major_tick_mark(self, value):
        self._element._remove_majorTickMark()
        if value is XL_TICK_MARK.CROSS:
            return
        self._element._add_majorTickMark(val=value)

    @property
    def maximum_scale(self):
        """
        Read/write float value specifying upper limit of value range, the
        number at the top of the vertical (value) scale. |None| if no maximum
        scale has been set.
        """
        return self._element.scaling.maximum

    @maximum_scale.setter
    def maximum_scale(self, value):
        scaling = self._element.scaling
        scaling.maximum = value

    @property
    def minimum_scale(self):
        """
        Read/write float value specifying lower limit of value range, the
        number at the bottom or left of the value scale. |None| if no minimum
        scale has been set.
        """
        return self._element.scaling.minimum

    @minimum_scale.setter
    def minimum_scale(self, value):
        scaling = self._element.scaling
        scaling.minimum = value

    @property
    def minor_tick_mark(self):
        """
        Read/write :ref:`XlTickMark` value specifying the type of minor tick
        mark for this axis.
        """
        minorTickMark = self._element.minorTickMark
        if minorTickMark is None:
            return XL_TICK_MARK.CROSS
        return minorTickMark.val

    @minor_tick_mark.setter
    def minor_tick_mark(self, value):
        self._element._remove_minorTickMark()
        if value is XL_TICK_MARK.CROSS:
            return
        self._element._add_minorTickMark(val=value)

    @lazyproperty
    def tick_labels(self):
        """
        The |TickLabels| instance providing access to axis tick label
        formatting properties.
        """
        return TickLabels(self._element)

    @property
    def visible(self):
        """
        Read/write. |True| if axis is visible, |False| otherwise.
        """
        delete = self._element.delete
        if delete is None:
            return False
        return False if delete.val else True

    @visible.setter
    def visible(self, value):
        if value not in (True, False):
            raise ValueError(
                "assigned value must be True or False, got: %s" % value
            )
        delete = self._element.get_or_add_delete()
        delete.val = not value


class CategoryAxis(_BaseAxis):
    """
    A category axis of a chart.
    """


class TickLabels(object):
    """
    A service class providing access to formatting of axis tick mark labels.
    """
    def __init__(self, xAx_elm):
        super(TickLabels, self).__init__()
        self._element = xAx_elm


class ValueAxis(_BaseAxis):
    """
    A value axis of a chart.
    """
