# encoding: utf-8

"""
Objects related to mouse click and hover actions on a shape or text.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .enum.action import PP_ACTION
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

    @property
    def action(self):
        """
        A member of the :ref:`PpActionType` enumeration, such as
        `PP_ACTION.HYPERLINK`, indicating the type of action that will result
        when the specified shape or text is clicked or the mouse pointer is
        positioned over the shape during a slide show.
        """
        if self._hover:
            hlink = self._element.hlinkHover
        else:
            hlink = self._element.hlinkClick

        if hlink is None:
            return PP_ACTION.NONE

        action_verb = hlink.action_verb

        if action_verb == 'hlinkshowjump':
            relative_target = hlink.action_fields['jump']
            return {
                'firstslide':      PP_ACTION.FIRST_SLIDE,
                'lastslide':       PP_ACTION.LAST_SLIDE,
                'lastslideviewed': PP_ACTION.LAST_SLIDE_VIEWED,
                'nextslide':       PP_ACTION.NEXT_SLIDE,
                'previousslide':   PP_ACTION.PREVIOUS_SLIDE,
                'endshow':         PP_ACTION.END_SHOW,
            }[relative_target]

        return {
            None:           PP_ACTION.HYPERLINK,
            'hlinksldjump': PP_ACTION.NAMED_SLIDE,
            'hlinkpres':    PP_ACTION.PLAY,
            'hlinkfile':    PP_ACTION.OPEN_FILE,
            'customshow':   PP_ACTION.NAMED_SLIDE_SHOW,
            'ole':          PP_ACTION.OLE_VERB,
            'macro':        PP_ACTION.RUN_MACRO,
            'program':      PP_ACTION.RUN_PROGRAM,
        }[action_verb]
