# encoding: utf-8

"""
Axis-related chart objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..enum.chart import XL_TICK_LABEL_POSITION, XL_TICK_MARK
from ..util import lazyproperty


class _BaseAxis(object):
    """
    Base class for chart axis classes.
    """
    def __init__(self, xAx_elm):
        super(_BaseAxis, self).__init__()
        self._element = xAx_elm

    @property
    def has_major_gridlines(self):
        """
        Read/write boolean value specifying whether this axis has gridlines
        at its major tick mark locations.
        """
        if self._element.majorGridlines is None:
            return False
        return True

    @has_major_gridlines.setter
    def has_major_gridlines(self, value):
        if bool(value) is True:
            self._element.get_or_add_majorGridlines()
        else:
            self._element._remove_majorGridlines()

    @property
    def has_minor_gridlines(self):
        """
        Read/write boolean value specifying whether this axis has gridlines
        at its minor tick mark locations.
        """
        if self._element.minorGridlines is None:
            return False
        return True

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
    def tick_label_position(self):
        """
        Read/write :ref:`XlTickLabelPosition` value specifying where the tick
        labels for this axis should appear.
        """
        tickLblPos = self._element.tickLblPos
        if tickLblPos is None:
            return XL_TICK_LABEL_POSITION.NEXT_TO_AXIS
        if tickLblPos.val is None:
            return XL_TICK_LABEL_POSITION.NEXT_TO_AXIS
        return tickLblPos.val

    @tick_label_position.setter
    def tick_label_position(self, value):
        tickLblPos = self._element.get_or_add_tickLblPos()
        tickLblPos.val = value

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

    @property
    def number_format(self):
        """
        Read/write string specifying the format for the numbers on this axis.
        Returns 'General' if no number format has been set. Note that this
        format string has no effect on rendered tick labels when
        :meth:`number_format_is_linked` is |True|. Assigning a format string
        to this property automatically sets :meth:`number_format_is_linked`
        to |False|.
        """
        numFmt = self._element.numFmt
        if numFmt is None:
            return 'General'
        return numFmt.formatCode

    @number_format.setter
    def number_format(self, value):
        numFmt = self._element.get_or_add_numFmt()
        numFmt.formatCode = value
        self.number_format_is_linked = False

    @property
    def number_format_is_linked(self):
        """
        Read/write boolean specifying whether number formatting should be
        taken from the source spreadsheet rather than the value of
        :meth:`number_format`.
        """
        numFmt = self._element.numFmt
        if numFmt is None:
            return False
        souceLinked = numFmt.sourceLinked
        if souceLinked is None:
            return True
        return numFmt.sourceLinked

    @number_format_is_linked.setter
    def number_format_is_linked(self, value):
        numFmt = self._element.get_or_add_numFmt()
        numFmt.sourceLinked = value


class ValueAxis(_BaseAxis):
    """
    A value axis of a chart.
    """
