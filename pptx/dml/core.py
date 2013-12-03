# encoding: utf-8

"""
Core DrawingML objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from types import NoneType

from pptx.enum import (
    MSO_COLOR_TYPE, MSO_FILL_TYPE as MSO_FILL, MSO_THEME_COLOR
)
from pptx.oxml.dml import (
    CT_BlipFillProperties, CT_GradientFillProperties, CT_GroupFillProperties,
    CT_HslColor, CT_NoFillProperties, CT_PatternFillProperties,
    CT_PresetColor, CT_SchemeColor, CT_ScRgbColor,
    CT_SolidColorFillProperties, CT_SRgbColor, CT_SystemColor
)
from pptx.util import lazyproperty


class ColorFormat(object):
    """
    Provides access to color settings such as RGB color, theme color, and
    luminance adjustments.
    """
    def __init__(self, eg_colorchoice_parent, color):
        super(ColorFormat, self).__init__()
        self._xFill = eg_colorchoice_parent
        self._color = color

    @property
    def brightness(self):
        """
        Read/write float value between -1.0 and 1.0 indicating the brightness
        adjustment for this color, e.g. -0.25 is 25% darker and 0.4 is 40%
        lighter. 0 means no brightness adjustment.
        """
        return self._color.brightness

    @classmethod
    def from_colorchoice_parent(cls, eg_colorchoice_parent):
        xClr = eg_colorchoice_parent.eg_colorchoice
        color = _Color(xClr)
        color_format = cls(eg_colorchoice_parent, color)
        return color_format

    @property
    def rgb(self):
        """
        |RGBColor| value of this color, or None if no RGB color is explicitly
        defined for this font. Setting this value to an |RGBColor| instance
        cause its type to change to MSO_COLOR_TYPE.RGB. If the color was a
        theme color with a brightness adjustment, the brightness adjustment
        is removed when changing it to an RGB color.
        """
        return self._color.rgb

    @rgb.setter
    def rgb(self, rgb):
        if not isinstance(rgb, RGBColor):
            raise ValueError('assigned value must be type RGBColor')
        # change to rgb color format if not already
        if not isinstance(self._color, _SRgbColor):
            srgbClr = self._xFill.get_or_change_to_srgbClr()
            self._color = _SRgbColor(srgbClr)
        # call _SRgbColor instance to do the setting
        self._color.rgb = rgb

    @property
    def theme_color(self):
        """
        Theme color value of this color, one of those defined in the
        MSO_THEME_COLOR enumeration, e.g. MSO_THEME_COLOR.ACCENT_1. Raises
        AttributeError on access if the color is not type
        ``MSO_COLOR_TYPE.SCHEME``. Assigning a value in ``MSO_THEME_COLOR``
        causes the color's type to change to ``MSO_COLOR_TYPE.SCHEME``.
        """
        return self._color.theme_color

    @property
    def type(self):
        """
        Read-only. A value from MSO_COLOR_TYPE, either RGB or SCHEME,
        corresponding to the way this color is defined, or None if no color
        is defined at the level of this font.
        """
        return self._color.color_type


class _Color(object):
    """
    Object factory for color object of the appropriate type, also the base
    class for all color type classes such as SRgbColor.
    """
    def __new__(cls, xClr):
        subcls_map = {
            NoneType:       _NoneColor,
            CT_HslColor:    _HslColor,
            CT_PresetColor: _PrstColor,
            CT_SchemeColor: _SchemeColor,
            CT_ScRgbColor:  _ScRgbColor,
            CT_SRgbColor:   _SRgbColor,
            CT_SystemColor: _SysColor,
        }
        color_cls = subcls_map[type(xClr)]
        return super(_Color, cls).__new__(color_cls)

    def __init__(self, xClr):
        super(_Color, self).__init__(xClr)
        self._xClr = xClr

    @property
    def brightness(self):
        from lxml import etree
        print(etree.tostring(self._xClr))
        print(type(self))
        lumMod, lumOff = self._xClr.lumMod, self._xClr.lumOff
        # a tint is lighter, a shade is darker
        # only tints have lumOff child
        if lumOff is not None:
            val = lumOff.val
            brightness = float(val) / 100000
            return brightness
        # which leaves shades, if lumMod is present
        if lumMod is not None:
            val = lumMod.val
            brightness = -1.0 + float(val)/100000
            return brightness
        # there's no brightness adjustment if no lum{Mod|Off} elements
        return 0

    @property
    def color_type(self):
        tmpl = ".color_type property must be implemented on %s"
        raise NotImplementedError(tmpl % self.__class__.__name__)

    @property
    def rgb(self):
        """
        Raises TypeError on access unless overridden by subclass.
        """
        tmpl = "no .rgb property on color type '%s'"
        raise AttributeError(tmpl % self.__class__.__name__)

    @property
    def theme_color(self):
        """
        Raises TypeError on access unless overridden by subclass.
        """
        tmpl = "no .theme_color property on color type '%s'"
        raise AttributeError(tmpl % self.__class__.__name__)


class _HslColor(_Color):

    @property
    def color_type(self):
        return MSO_COLOR_TYPE.HSL


class _NoneColor(_Color):

    @property
    def color_type(self):
        return None


class _PrstColor(_Color):

    @property
    def color_type(self):
        return MSO_COLOR_TYPE.PRESET


class _SchemeColor(_Color):

    def __init__(self, schemeClr):
        super(_SchemeColor, self).__init__(schemeClr)
        self._schemeClr = schemeClr

    @property
    def color_type(self):
        return MSO_COLOR_TYPE.SCHEME

    @property
    def theme_color(self):
        """
        Theme color value of this color, one of those defined in the
        MSO_THEME_COLOR enumeration, e.g. MSO_THEME_COLOR.ACCENT_1. None if
        no theme color is explicitly defined for this font. Setting this to a
        value in MSO_THEME_COLOR causes the color's type to change to
        ``MSO_COLOR_TYPE.SCHEME``.
        """
        return MSO_THEME_COLOR.from_xml(self._schemeClr.val)


class _ScRgbColor(_Color):

    @property
    def color_type(self):
        return MSO_COLOR_TYPE.SCRGB


class _SRgbColor(_Color):

    def __init__(self, srgbClr):
        super(_SRgbColor, self).__init__(srgbClr)
        self._srgbClr = srgbClr

    @property
    def color_type(self):
        return MSO_COLOR_TYPE.RGB

    @property
    def rgb(self):
        """
        |RGBColor| value of this color, corresponding to the value in the
        required ``val`` attribute of the ``<a:srgbColr>`` element.
        """
        return RGBColor.from_string(self._srgbClr.val)

    @rgb.setter
    def rgb(self, rgb):
        self._srgbClr.val = str(rgb)


class _SysColor(_Color):

    @property
    def color_type(self):
        return MSO_COLOR_TYPE.SYSTEM


class FillFormat(object):
    """
    Provides access to the current fill properties object and provides
    methods to change the fill type.
    """
    def __init__(self, eg_fillproperties_parent, fill_obj):
        super(FillFormat, self).__init__()
        self._xPr = eg_fillproperties_parent
        self._fill = fill_obj

    @property
    def fill_type(self):
        """
        Return a |ColorFormat| instance representing the foreground color of
        this fill.
        """
        return self._fill.fill_type

    @property
    def fore_color(self):
        """
        Return a |ColorFormat| instance representing the foreground color of
        this fill.
        """
        return self._fill.fore_color

    @classmethod
    def from_fill_parent(cls, eg_fillproperties_parent):
        """
        Return a |FillFormat| instance initialized to the settings contained
        in *eg_fill_properties_parent*, which must be an element having
        EG_FillProperties in its schema sequence.
        """
        fill_elm = eg_fillproperties_parent.eg_fillproperties
        fill = _Fill(fill_elm)
        fill_format = cls(eg_fillproperties_parent, fill)
        return fill_format

    def solid(self):
        """
        Sets the fill type to solid, i.e. a solid color. Note that calling
        this method does not set a color or by itself cause the shape to
        appear with a solid color fill; rather it enables subsequent
        assignments to properties like fore_color to set the color.
        """
        solidFill = self._xPr.get_or_change_to_solidFill()
        self._fill = _SolidFill(solidFill)


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


class _Fill(object):
    """
    Object factory for fill object of class matching fill element, such as
    _SolidFill for ``<a:solidFill>``; also serves as the base class for all
    fill classes
    """
    def __new__(cls, xFill):
        if xFill is None:
            fill_cls = _NoneFill
        elif isinstance(xFill, CT_BlipFillProperties):
            fill_cls = _BlipFill
        elif isinstance(xFill, CT_GradientFillProperties):
            fill_cls = _GradFill
        elif isinstance(xFill, CT_GroupFillProperties):
            fill_cls = _GrpFill
        elif isinstance(xFill, CT_NoFillProperties):
            fill_cls = _NoFill
        elif isinstance(xFill, CT_PatternFillProperties):
            fill_cls = _PattFill
        elif isinstance(xFill, CT_SolidColorFillProperties):
            fill_cls = _SolidFill
        else:
            fill_cls = _Fill
        return super(_Fill, cls).__new__(fill_cls)

    @property
    def fore_color(self):
        """
        Raise NotImplementedError for all fill types that are still skeleton
        subclasses.
        """
        tmpl = ".fore_color property not implemented yet for %s"
        raise NotImplementedError(tmpl % self.__class__.__name__)

    @property
    def fill_type(self):
        tmpl = ".fill_type property must be implemented on %s"
        raise NotImplementedError(tmpl % self.__class__.__name__)


class _BlipFill(_Fill):

    @property
    def fill_type(self):
        return MSO_FILL.PICTURE

    @property
    def fore_color(self):
        """
        Raise TypeError with message explaining why this doesn't make sense.
        """
        tmpl = "a picture fill has no foreground color"
        raise TypeError(tmpl)


class _GradFill(_Fill):

    @property
    def fill_type(self):
        return MSO_FILL.GRADIENT


class _GrpFill(_Fill):

    @property
    def fill_type(self):
        return MSO_FILL.GROUP

    @property
    def fore_color(self):
        """
        Raise TypeError with message explaining why this doesn't make sense.
        """
        tmpl = "a group fill has no foreground color"
        raise TypeError(tmpl)


class _NoFill(_Fill):

    @property
    def fill_type(self):
        return MSO_FILL.BACKGROUND


class _NoneFill(_Fill):

    @property
    def fill_type(self):
        return None

    @property
    def fore_color(self):
        """
        Raise TypeError with message explaining why this doesn't make sense.
        """
        tmpl = "can't set .fore_color on no fill, call .solid() first"
        raise TypeError(tmpl)


class _PattFill(_Fill):

    @property
    def fill_type(self):
        return MSO_FILL.PATTERNED


class _SolidFill(_Fill):
    """
    Provides access to fill properties such as color for solid fills.
    """
    def __init__(self, solidFill):
        super(_SolidFill, self).__init__()
        self._solidFill = solidFill

    @property
    def fill_type(self):
        return MSO_FILL.SOLID

    @lazyproperty
    def fore_color(self):
        return ColorFormat.from_colorchoice_parent(self._solidFill)
