# encoding: utf-8

"""
|ChartFormat| and related objects. |ChartFormat| acts as proxy for the `spPr`
element, which provides visual shape properties such as line and fill for
chart elements.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..shared import ElementProxy


class ChartFormat(ElementProxy):
    """
    Provides access to visual shape properties such as line and fill.
    """

    __slots__ = ()
