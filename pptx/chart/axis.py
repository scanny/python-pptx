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


class ValueAxis(_BaseAxis):
    """
    A value axis of a chart.
    """
