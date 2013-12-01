# encoding: utf-8

"""
Core DrawingML objects.
"""


class ColorFormat(object):
    """
    Provides access to color settings such as RGB color, theme color, and
    luminance adjustments.
    """


class FillFormat(object):
    """
    Provides access to the current fill properties object and provides
    methods to change the fill type.
    """


class RGBColor(tuple):
    """
    Immutable value object defining a particular RGB color.
    """
    def __new__(cls, r, g, b):
        msg = 'RGBColor() takes three integer values 0-255'
        for val in (r, g, b):
            if not isinstance(val, int) or val < 0 or val > 255:
                raise ValueError(msg)
        return super(RGBColor, cls).__new__(cls, (r, g, b))

    def __str__(self):
        """
        Return a hex string rgb value, like '3C2F80'
        """
        return '%02X%02X%02X' % self

    @classmethod
    def from_string(cls, rgb_hex_str):
        r = int(rgb_hex_str[:2], 16)
        g = int(rgb_hex_str[2:4], 16)
        b = int(rgb_hex_str[4:], 16)
        return cls(r, g, b)
