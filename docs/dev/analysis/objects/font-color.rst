##########
Font Color
##########

:Updated:  2013-11-07
:Author:   Steve Canny
:Status:   **WORKING DRAFT**


Minimal implementation::

    font.color = RGBColor(0x3F, 0x2c, 0x36)

Draft API
=========

RGBColor would be an immutable value object that could be reused as often as
needed and not tied to any part of the underlying XML tree.

::

    font.color.rgb = RGB(0x3F, 0x2c, 0x36)
    font.color.theme_color = MSO_THEME_COLOR_ACCENT_1

_Font.color is a ColorFormat subclass, perhaps FontColor since it would have
font-specific behaviors::

    class _Font(...):
        @property
        def color(self):
            if not hasattr(self, _color):
                self._color = _FontColor(self._rPr)
            return self._color

    class _FontColor(object):
        def __init__(self, rPr):
            self._rPr = rPr

        def clear(self):
            self._rPr.remove_fill()

        @property
        def rgb(self):
            srgbClr = self._rPr.srgbClr
            if srgbClr = None
                return None
            rgb = RGB.from_str(srgbClr.rgb)
            return rgb

        @rgb.setter
        def rgb(self, rgb):
            solidFill = self._rPr.get_or_change_to_solidFill()
            srgbClr = solidFill.get_or_change_to_srgbClr()
            srgbClr.rgb = rgb

        @property
        def theme_color(self):
            schemeClr = self._rPr.schemeClr
            if schemeClr = None
                return None
            theme_color = MsoThemeColorIndex.from_xml(schemeClr.val)
            return rgb

        @rgb.setter
        def rgb(self, rgb):
            solidFill = self._rPr.get_or_change_to_solidFill()
            srgbClr = solidFill.get_or_change_to_srgbClr()
            srgbClr.rgb = rgb

    class RGB(object):
        def __init__(self, r, g, b):
            self._r = r
            self._g = g
            self._b = b
        @property
        def __str__(self):
            return '%02x%02x%02x' % (self._r, self._g, self._b)

    def _get_color(self):
        if self.fill is None:
            return None
        return self.fill.fore_color

_Font.color setter is a shortcut for::

    def _set_color(self, color):
        self.fill.solid()
        self.fill.fore_color.rgb = RGB(0, 128, 128)
        self.fill.fore_color.theme_color = MSO_THEME_COLOR_ACCENT_1


Notes on fill API
=================

_Font.fill getter returns a subclass of type Fill, or None
_Font.fill setter accepts a subclass of type Fill, or None

fill.type is overridden by each subclass to return the right type enumeration
constant

This implies::

    module pptx.dml.fill

    class _BaseFill(object):
        pass

    class SolidFill(_BaseFill):

        @accepts(color=(None, _BaseColor))
        def __init__(self, color=None):
            self._color = color

        @property
        def fore_color(self):
            return self._color

        @property
        def type(self):
            return MSO_FILL_SOLID

    class _BaseColor(object):
        pass

    class RGBColor(_BaseColor):

        def __init__(self, r, g, b):
            ...
