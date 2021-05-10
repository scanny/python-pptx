# encoding: utf-8

"""Theme-releated objects including theme, color, font, and format schemes."""

from __future__ import absolute_import, division, print_function, unicode_literals

from pptx.shared import ElementProxy
from pptx.dml.color import ColorFormat
from pptx.dml.fill import FillFormat, _Fill
from pptx.dml.line import LineStyle
from pptx.text.text import TextFont


class Theme(ElementProxy):
    """Provides access to the Theme associated with a MasterSlide

    This is a read only object.
    """
    @property
    def name(self):
        """ Return read only access to the :attr: `name`
        """
        return self._element.name

    @property
    def theme_elements(self):
        return ThemeElements(self._element.theme_elements)

    @property
    def color_scheme(self):
        """Easy access to theme element's color_scheme"""

        return self.theme_elements.color_scheme
    
    @property
    def font_scheme(self):
        """Easy access to theme elements' font_scheme"""
        return self.theme_elements.font_scheme
        
    @property
    def format_scheme(self):
        """Easy access to theme elements' format_scheme"""
        return self.theme_elements.format_scheme
        

class ThemeElements(ElementProxy):
    """Provides access to the theme elements"""

    @property
    def color_scheme(self):
        """Read Only Access to Theme's Color Scheme"""
        return ColorScheme(self._element.clrScheme)
    
    @property
    def font_scheme(self):
        """ Read Only Access to Theme's Font Scheme"""
        return FontScheme(self._element.fontScheme)

    @property
    def format_scheme(self):
        """Read only access to theme's format scheme"""
        return FormatScheme(self._element.fmtScheme)


class ColorScheme(ElementProxy):
    @property
    def name(self):
        return self._element.name

    @property
    def dk1(self):
        return ColorFormat.from_colorchoice_parent(self._element.dk1)

    @property
    def dk2(self):
        return ColorFormat.from_colorchoice_parent(self._element.dk2)

    @property
    def lt1(self):
        return ColorFormat.from_colorchoice_parent(self._element.lt1)

    @property
    def lt2(self):
        return ColorFormat.from_colorchoice_parent(self._element.lt2)

    @property
    def accent1(self):
        return ColorFormat.from_colorchoice_parent(self._element.accent1)

    @property
    def accent2(self):
        return ColorFormat.from_colorchoice_parent(self._element.accent2)

    @property
    def accent3(self):
        return ColorFormat.from_colorchoice_parent(self._element.accent3)

    @property
    def accent4(self):
        return ColorFormat.from_colorchoice_parent(self._element.accent4)

    @property
    def accent5(self):
        return ColorFormat.from_colorchoice_parent(self._element.accent5)

    @property
    def accent6(self):
        return ColorFormat.from_colorchoice_parent(self._element.accent6)

    @property
    def hyperlink(self):
        """ Plain English propery name for `hlink`"""
        return self.hlink

    @property
    def followed_hyperlink(self):
        """ Plain English propery name for `folHlink`"""
        return self.folHlink

    @property
    def hlink(self):
        return ColorFormat.from_colorchoice_parent(self._element.hlink)

    @property
    def folHlink(self):
        return ColorFormat.from_colorchoice_parent(self._element.folHlink)


class FontScheme(ElementProxy):
    @property
    def major_font(self):
        return FontCollection(self._element.majorFont)

    @property
    def minor_font(self):
        return FontCollection(self._element.minorFont)


class FormatScheme(ElementProxy):
    @property
    def name(self):
        return self._element.name

    @property
    def fill_style_list(self):
        return tuple([FillFormat(self._element.fillStyleLst, _Fill(fill)) for fill in self._element.fillStyleLst])
        
    @property
    def line_style_list(self):
        return tuple([LineStyle(line) for line in self._element.lnStyleLst])
    
    @property
    def effect_style_list(self):
        return []

    @property
    def background_fill_style_list(self):
        return tuple([FillFormat(self._element.bgFillStyleLst, _Fill(fill)) for fill in self._element.bgFillStyleLst])


class FontCollection(ElementProxy):
    @property
    def latin(self): 
        return TextFont(self.element.latin)

    @property
    def east_asian(self): 
        return TextFont(self.element.ea)

    @property
    def complex_script(self): 
        return TextFont(self.element.cs)

class ColorMap(ElementProxy):
    @property
    def bg1(self):
        return self.element.bg1

    @property
    def tx1(self):
        return self.element.tx1
    
    @property
    def bg2(self):
        return self.element.bg2
    
    @property
    def tx2(self):
        return self.element.tx2
    
    @property
    def accent1(self):
        return self.element.accent1

    @property
    def accent2(self):
        return self.element.accent2

    @property
    def accent3(self):
        return self.element.accent3

    @property
    def accent4(self):
        return self.element.accent4

    @property
    def accent5(self):
        return self.element.accent5

    @property
    def accent6(self):
        return self.element.accent6

    @property
    def hyperlink(self):
        return self.element.hlink

    @property
    def followed_hyperlink(self):
        return self.element.folHlink

