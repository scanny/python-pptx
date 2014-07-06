# encoding: utf-8

"""
Enumerations used by charts and related objects
"""

from __future__ import absolute_import

from .base import XmlEnumeration, XmlMappedEnumMember


class XL_TICK_MARK(XmlEnumeration):
    """
    Specifies a type of axis tick for a chart.

    Example::

        from pptx.enum.chart import XL_TICK_MARK

        chart.value_axis.minor_tick_mark = XL_TICK_MARK.INSIDE
    """

    __ms_name__ = 'XlTickMark'

    __url__ = (
        'http://msdn.microsoft.com/en-us/library/office/ff193878.aspx'
    )

    __members__ = (
        XmlMappedEnumMember(
            'CROSS', 4, 'cross', 'Tick mark crosses the axis'
        ),
        XmlMappedEnumMember(
            'INSIDE', 2, 'in', 'Tick mark appears inside the axis'
        ),
        XmlMappedEnumMember(
            'NONE', -4142, 'none', 'No tick mark'
        ),
        XmlMappedEnumMember(
            'OUTSIDE', 3, 'out', 'Tick mark appears outside the axis'
        ),
    )


class XL_TICK_LABEL_POSITION(XmlEnumeration):
    """
    Specifies the position of tick-mark labels on a chart axis.

    Example::

        from pptx.enum.chart import XL_TICK_LABEL_POSITION

        category_axis = chart.category_axis
        category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    """

    __ms_name__ = 'XlTickLabelPosition'

    __url__ = (
        'http://msdn.microsoft.com/en-us/library/office/ff822561.aspx'
    )

    __members__ = (
        XmlMappedEnumMember(
            'HIGH', -4127, 'high', 'Top or right side of the chart.'
        ),
        XmlMappedEnumMember(
            'LOW', -4134, 'low', 'Bottom or left side of the chart.'
        ),
        XmlMappedEnumMember(
            'NEXT_TO_AXIS', 4, 'nextTo', 'Next to axis (where axis is not at'
            ' either side of the chart).'
        ),
        XmlMappedEnumMember(
            'NONE', -4142, 'none', 'No tick labels.'
        ),
    )
