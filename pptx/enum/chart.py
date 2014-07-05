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
