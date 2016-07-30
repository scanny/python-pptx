# encoding: utf-8

"""
Marker-related objects. Only the line-type charts Line, XY, and Radar have
markers.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..shared import ElementProxy


class Marker(ElementProxy):
    """
    Represents a data point marker, such as a diamond or circle, on
    a line-type chart.
    """

    __slots__ = ()
