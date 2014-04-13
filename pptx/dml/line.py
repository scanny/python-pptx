# encoding: utf-8

"""
DrawingML objects related to line formatting
"""

from __future__ import absolute_import, print_function, unicode_literals

from .fill import FillFormat
from ..util import lazyproperty


class LineFormat(object):
    """
    Provides access to line properties such as line color, style, and width.
    """
    def __init__(self, parent):
        super(LineFormat, self).__init__()
        self._parent = parent

    @lazyproperty
    def fill(self):
        """
        |FillFormat| instance for this line, providing access to fill
        properties such as foreground color.
        """
        ln = self._get_or_add_ln()
        return FillFormat.from_fill_parent(ln)

    def _get_or_add_ln(self):
        """
        Return the ``<a:ln>`` element containing the line format properties
        in the XML.
        """
        return self._parent.get_or_add_ln()
