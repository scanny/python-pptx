# encoding: utf-8

"""
Enumerations used by DrawingML objects
"""

from __future__ import absolute_import

from .base import (
    alias, Enumeration, EnumMember, ReturnValueOnlyEnumMember,
    XmlEnumeration, XmlMappedEnumMember
)


class MSO_COLOR_TYPE(Enumeration):
    """
    Specifies the color specification scheme

    Example::

        from pptx.enum.dml import MSO_COLOR_TYPE

        assert shape.fill.fore_color.type == MSO_COLOR_TYPE.SCHEME
    """

    __ms_name__ = 'MsoColorType'

    __url__ = (
        'http://msdn.microsoft.com/en-us/library/office/ff864912(v=office.15'
        ').aspx'
    )

    __members__ = (
        EnumMember(
            'RGB', 1, 'Color is specified by an |RGBColor| value'
        ),
        EnumMember(
            'SCHEME', 2, 'Color is one of the preset theme colors'
        ),
        EnumMember(
            'HSL', 101, """
            Color is specified using Hue, Saturation, and Luminosity values
            """
        ),
        EnumMember(
            'PRESET', 102, """
            Color is specified using a named built-in color
            """
        ),
        EnumMember(
            'SCRGB', 103, """
            Color is an scRGB color, a wide color gamut RGB color space
            """
        ),
        EnumMember(
            'SYSTEM', 104, """
            Color is one specified by the operating system, such as the
            window background color.
            """
        ),
    )


@alias('MSO_FILL')
class MSO_FILL_TYPE(Enumeration):
    """
    Specifies the type of bitmap used for the fill of a shape.

    Alias: ``MSO_FILL``

    Example::

        from pptx.enum.dml import MSO_FILL

        assert shape.fill.type == MSO_FILL.SOLID
    """

    __ms_name__ = 'MsoFillType'

    __url__ = 'http://msdn.microsoft.com/EN-US/library/office/ff861408.aspx'

    __members__ = (
        EnumMember(
            'BACKGROUND', 5, """
            The shape is transparent, such that whatever is behind the shape
            shows through. Often this is the slide background, but if
            a visible shape is behind, that will show through.
            """
        ),
        EnumMember(
            'GRADIENT', 3, 'Shape is filled with a gradient'
        ),
        EnumMember(
            'GROUP', 101, 'Shape is part of a group and should inherit the '
            'fill properties of the group.'
        ),
        EnumMember(
            'PATTERNED', 2, 'Shape is filled with a pattern'
        ),
        EnumMember(
            'PICTURE', 6, 'Shape is filled with a bitmapped image'
        ),
        EnumMember(
            'SOLID', 1, 'Shape is filled with a solid color'
        ),
        EnumMember(
            'TEXTURED', 4, 'Shape is filled with a texture'
        ),
    )


@alias('MSO_THEME_COLOR')
class MSO_THEME_COLOR_INDEX(XmlEnumeration):
    """
    Indicates the Office theme color, one of those shown in the color gallery
    on the formatting ribbon.

    Alias: ``MSO_THEME_COLOR``

    Example::

        from pptx.enum.dml import MSO_THEME_COLOR

        shape.fill.solid()
        shape.fill.fore_color.theme_color == MSO_THEME_COLOR.ACCENT_1
    """

    __ms_name__ = 'MsoThemeColorIndex'

    __url__ = (
        'http://msdn.microsoft.com/en-us/library/office/ff860782(v=office.15'
        ').aspx'
    )

    __members__ = (
        EnumMember(
            'NOT_THEME_COLOR', 0, 'Indicates the color is not a theme color.'
        ),
        XmlMappedEnumMember(
            'ACCENT_1', 5, 'accent1', 'Specifies the Accent 1 theme color.'
        ),
        XmlMappedEnumMember(
            'ACCENT_2', 6, 'accent2', 'Specifies the Accent 2 theme color.'
        ),
        XmlMappedEnumMember(
            'ACCENT_3', 7, 'accent3', 'Specifies the Accent 3 theme color.'
        ),
        XmlMappedEnumMember(
            'ACCENT_4', 8, 'accent4', 'Specifies the Accent 4 theme color.'
        ),
        XmlMappedEnumMember(
            'ACCENT_5', 9, 'accent5', 'Specifies the Accent 5 theme color.'
        ),
        XmlMappedEnumMember(
            'ACCENT_6', 10, 'accent6', 'Specifies the Accent 6 theme color.'
        ),
        XmlMappedEnumMember(
            'BACKGROUND_1', 14, 'bg1', 'Specifies the Background 1 theme '
            'color.'
        ),
        XmlMappedEnumMember(
            'BACKGROUND_2', 16, 'bg2', 'Specifies the Background 2 theme '
            'color.'
        ),
        XmlMappedEnumMember(
            'DARK_1', 1, 'dk1', 'Specifies the Dark 1 theme color.'
        ),
        XmlMappedEnumMember(
            'DARK_2', 3, 'dk2', 'Specifies the Dark 2 theme color.'
        ),
        XmlMappedEnumMember(
            'FOLLOWED_HYPERLINK', 12, 'folHlink', 'Specifies the theme color'
            ' for a clicked hyperlink.'
        ),
        XmlMappedEnumMember(
            'HYPERLINK', 11, 'hlink', 'Specifies the theme color for a hyper'
            'link.'
        ),
        XmlMappedEnumMember(
            'LIGHT_1', 2, 'lt1', 'Specifies the Light 1 theme color.'
        ),
        XmlMappedEnumMember(
            'LIGHT_2', 4, 'lt2', 'Specifies the Light 2 theme color.'
        ),
        XmlMappedEnumMember(
            'TEXT_1', 13, 'tx1', 'Specifies the Text 1 theme color.'
        ),
        XmlMappedEnumMember(
            'TEXT_2', 15, 'tx2', 'Specifies the Text 2 theme color.'
        ),
        ReturnValueOnlyEnumMember(
            'MIXED', -2, 'Indicates multiple theme colors are used, such as '
            'in a group shape.'
        ),
    )
