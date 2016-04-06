# encoding: utf-8

"""
Chart title.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..text.text import TextFrame


class ChartTitle(object):
    """
    Represents a chart title. A chart or axis can have at most one title.
    """
    def __init__(self, title_elm):
        super(ChartTitle, self).__init__()
        self._element = title_elm

    @property
    def text_frame(self):
        """
        |TextFrame| instance for this shape, containing the text of the shape
        and providing access to text formatting properties.
        """
        return TextFrame(self._element.text_frame, self._element.tx)
