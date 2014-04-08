# encoding: utf-8

"""
Enumerations used by text and related objects
"""

from __future__ import absolute_import

from .base import Enumeration, EnumMember, ReturnValueOnlyEnumMember


class MSO_AUTO_SIZE(Enumeration):
    """
    Determines the type of automatic sizing allowed.

    The following names can be used to specify the automatic sizing behavior
    used to fit a shape's text within the shape bounding box, for example::

        from pptx.enum.text import MSO_AUTO_SIZE

        shape.textframe.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    The word-wrap setting of the textframe interacts with the auto-size setting
    to determine the specific auto-sizing behavior.

    Note that ``TextFrame.auto_size`` can also be set to |None|, which removes
    the auto size setting altogether. This causes the setting to be inherited,
    either from the layout placeholder, in the case of a placeholder shape, or
    from the theme.
    """
    """
    Corresponds to MsoAutoSize enumeration
    http://msdn.microsoft.com/en-us/library/office/ff865367(v=office.15).aspx
    """
    NONE = 0
    SHAPE_TO_FIT_TEXT = 1
    TEXT_TO_FIT_SHAPE = 2

    __ms_name__ = 'MsoAutoSize'

    __url__ = (
        'http://msdn.microsoft.com/en-us/library/office/ff865367(v=office.15'
        ').aspx'
    )

    __members__ = (
        EnumMember(
            'NONE', 0, 'No automatic sizing of the shape or text will be don'
            'e. Text can freely extend beyond the horizontal and vertical ed'
            'ges of the shape bounding box.'
        ),
        EnumMember(
            'SHAPE_TO_FIT_TEXT', 1, 'The shape height and possibly width are'
            ' adjusted to fit the text. Note this setting interacts with the'
            ' TextFrame.word_wrap property setting. If word wrap is turned o'
            'n, only the height of the shape will be adjusted; soft line bre'
            'aks will be used to fit the text horizontally.'
        ),
        EnumMember(
            'TEXT_TO_FIT_SHAPE', 2, 'The font size is reduced as necessary t'
            'o fit the text within the shape.'
        ),
        ReturnValueOnlyEnumMember(
            'MIXED', -2, 'Return value only; indicates a combination of auto'
            'matic sizing schemes are used.'
        ),
    )
