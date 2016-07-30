# encoding: utf-8

"""
Marker-related objects. Only the line-type charts Line, XY, and Radar have
markers.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..dml.chtfmt import ChartFormat
from ..shared import ElementProxy
from ..util import lazyproperty


class Marker(ElementProxy):
    """
    Represents a data point marker, such as a diamond or circle, on
    a line-type chart.
    """

    __slots__ = ('_format',)

    @lazyproperty
    def format(self):
        """
        The |ChartFormat| instance for this marker, providing access to shape
        properties such as fill and line.
        """
        marker = self._element.get_or_add_marker()
        return ChartFormat(marker)
