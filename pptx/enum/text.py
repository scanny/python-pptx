# encoding: utf-8

"""
Enumerations used by text and related objects
"""

from __future__ import absolute_import

from .base import (
    alias, Enumeration, EnumMember, ReturnValueOnlyEnumMember,
    XmlEnumeration, XmlMappedEnumMember
)


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


@alias('MSO_ANCHOR')
class MSO_VERTICAL_ANCHOR(XmlEnumeration):
    """
    Specifies the vertical alignment of text in a text frame. Used with the
    ``.vertical_anchor`` property of the |TextFrame| object. Note that the
    ``vertical_anchor`` property can also have the value None, indicating
    there is no directly specified vertical anchor setting and its effective
    value is inherited from its placeholder if it has one or from the theme.
    None may also be assigned to remove an explicitly specified vertical
    anchor setting.
    """

    __ms_name__ = 'MsoVerticalAnchor'

    __url__ = 'http://msdn.microsoft.com/en-us/library/office/ff865255.aspx'

    __members__ = (
        XmlMappedEnumMember(
            None, None, None, 'Text frame has no vertical anchor specified '
            'and inherits its value from its layout placeholder or theme.'
        ),
        XmlMappedEnumMember(
            'TOP', 1, 't', 'Aligns text to top of text frame'
        ),
        XmlMappedEnumMember(
            'MIDDLE', 3, 'ctr', 'Centers text vertically'
        ),
        XmlMappedEnumMember(
            'BOTTOM', 4, 'b', 'Aligns text to bottom of text frame'
        ),
        ReturnValueOnlyEnumMember(
            'MIXED', -2, 'Return value only; indicates a combination of the '
            'other states.'
        ),
    )


@alias('PP_ALIGN')
class PP_PARAGRAPH_ALIGNMENT(XmlEnumeration):
    """
    Specifies the horizontal alignment for one or more paragraphs.

    Alias: ``PP_ALIGN``

    Example::

        from pptx.enum.text import PP_ALIGN

        shape.paragraphs[0].alignment = PP_ALIGN.CENTER
    """

    __ms_name__ = 'PpParagraphAlignment'

    __url__ = (
        'http://msdn.microsoft.com/en-us/library/office/ff745375(v=office.15'
        ').aspx'
    )

    __members__ = (
        XmlMappedEnumMember(
            None, None, None, 'No alignment setting on paragraph element'
        ),
        XmlMappedEnumMember(
            'CENTER', 2, 'ctr', 'Center align'
        ),
        XmlMappedEnumMember(
            'DISTRIBUTE', 5, 'dist', 'Evenly distributes e.g. Japanese chara'
            'cters from left to right within a line'
        ),
        XmlMappedEnumMember(
            'JUSTIFY', 4, 'just', 'Justified, i.e. each line both begins and'
            ' ends at the margin with spacing between words adjusted such th'
            'at the line exactly fills the width of the paragraph.'
        ),
        XmlMappedEnumMember(
            'JUSTIFY_LOW', 7, 'justLow', 'Justify using a small amount of sp'
            'ace between words.'
        ),
        XmlMappedEnumMember(
            'LEFT', 1, 'l', 'Left aligned'
        ),
        XmlMappedEnumMember(
            'RIGHT', 3, 'r', 'Right aligned'
        ),
        XmlMappedEnumMember(
            'THAI_DISTRIBUTE', 6, 'thaiDist', 'Thai distributed'
        ),
        ReturnValueOnlyEnumMember(
            'MIXED', -2, 'Return value only; indicates multiple paragraph al'
            'ignments are present in a set of paragraphs.'
        ),
    )
