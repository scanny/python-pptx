# encoding: utf-8

"""
DrawingML objects related to line formatting
"""

from __future__ import absolute_import, print_function, unicode_literals


class LineFormat(object):
    """
    Provides access to line properties such as line color, style, and width.
    """
    def __init__(self, parent):
        super(LineFormat, self).__init__()
        self._parent = parent
