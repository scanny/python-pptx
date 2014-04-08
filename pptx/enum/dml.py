# encoding: utf-8

"""
Enumerations used by DrawingML objects
"""

from __future__ import absolute_import

from .base import alias, Enumeration, EnumMember


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
