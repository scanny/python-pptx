# encoding: utf-8

"""
Enumerations used by DrawingML objects
"""

from __future__ import absolute_import

from .base import Enumeration, EnumMember


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
