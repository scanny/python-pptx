# encoding: utf-8

"""
Objects related to mouse click and hover actions on a shape or text.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .shapes import Subshape


class ActionSetting(Subshape):
    """
    Properties that specify how a shape or text run reacts to mouse actions
    during a slide show.
    """
    # Subshape superclass provides access to the Slide Part, which is needed
    # to access relationships.
    def __init__(self, xPr, parent, hover=False):
        super(ActionSetting, self).__init__(parent)
        # xPr is either a cNvPr or rPr element
        self._element = xPr
        # _hover determines use of `a:hlinkClick` or `a:hlinkHover`
        self._hover = hover
