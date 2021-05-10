# encoding: utf-8

"""
Test suite for pptx.theme module
"""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from pptx.theme import (
    Theme,
    ThemeElements,
    ColorScheme, 
    FontScheme,
    FormatScheme,
    FontCollection,
    ColorMap,
)
from pptx.dml.color import ColorFormat
from pptx.dml.fill import FillFormat
from pptx.dml.line import LineStyle
from pptx.text.text import TextFont

from .unitutil.cxml import element



class DescribeTheme:
    def it_knows_its_name(self, name_fixture):
        theme, expected_value = name_fixture
        assert theme.name == expected_value

    def it_gives_access_to_its_theme_elements(self, theme):
        assert isinstance(theme.theme_elements, ThemeElements)

    def it_gives_access_to_color_scheme(self, theme):
        assert isinstance(theme.color_scheme, ColorScheme)

    def it_gives_access_to_font_scheme(self, theme):
        assert isinstance(theme.font_scheme, FontScheme)

    def it_gives_Access_to_format_scheme(self, theme):
        assert isinstance(theme.format_scheme, FormatScheme)

    # fixtures ---------------------------------------------

    @pytest.fixture(
        params=[
            ("a:theme", ""),
            ("a:theme{name=Foobar}", "Foobar")
        ]
    )
    def name_fixture(self, request):
        theme_cxml, expected_value = request.param
        theme = Theme(element(theme_cxml))
        return theme, expected_value


    # fixture components ---------------------------------------------

    @pytest.fixture
    def theme(self):
        return Theme(element("a:theme/a:themeElements/(a:clrScheme,a:fontScheme,a:fmtScheme)"))

class DescribeThemeElements:
    def it_gives_access_to_color_scheme(self, theme_elements):
        assert isinstance(theme_elements.color_scheme, ColorScheme)

    def it_gives_access_to_font_scheme(self, theme_elements):
        assert isinstance(theme_elements.font_scheme, FontScheme)

    def it_gives_Access_to_format_scheme(self, theme_elements):
        assert isinstance(theme_elements.format_scheme, FormatScheme)

    # fixtures ---------------------------------------------

    # fixture components ---------------------------------------------

    @pytest.fixture
    def theme_elements(self):
        return ThemeElements(element("a:themeElements/(a:clrScheme,a:fontScheme,a:fmtScheme)"))

class DescribeColorScheme:
    def it_knows_its_name(self, name_fixture):
        color_scheme, expected_value = name_fixture
        assert color_scheme.name == expected_value
        
    def it_gives_access_to_dk1_color(self, color_scheme):
        assert isinstance(color_scheme.dk1, ColorFormat)

    def it_gives_access_to_dk2_color(self, color_scheme):
        assert isinstance(color_scheme.dk2, ColorFormat)

    def it_gives_access_to_lt1_color(self, color_scheme):
        assert isinstance(color_scheme.lt1, ColorFormat)

    def it_gives_access_to_lt2_color(self, color_scheme):
        assert isinstance(color_scheme.lt2, ColorFormat)

    def it_gives_access_to_accent1_color(self, color_scheme):
        assert isinstance(color_scheme.accent1, ColorFormat)

    def it_gives_access_to_accent2_color(self, color_scheme):
        assert isinstance(color_scheme.accent2, ColorFormat)

    def it_gives_access_to_accent3_color(self, color_scheme):
        assert isinstance(color_scheme.accent3, ColorFormat)

    def it_gives_access_to_accent4_color(self, color_scheme):
        assert isinstance(color_scheme.accent4, ColorFormat)

    def it_gives_access_to_accent5_color(self, color_scheme):
        assert isinstance(color_scheme.accent5, ColorFormat)

    def it_gives_access_to_accent6_color(self, color_scheme):
        assert isinstance(color_scheme.accent6, ColorFormat)

    def it_gives_access_to_hyperlink_color(self, color_scheme):
        assert isinstance(color_scheme.hlink, ColorFormat)

    def it_gives_access_to_hyperlink_color_lazy(self, color_scheme):
        assert isinstance(color_scheme.hyperlink, ColorFormat)

    def it_gives_access_to_followed_hyperlink_color_lazy(self, color_scheme):
        assert isinstance(color_scheme.followed_hyperlink, ColorFormat)

    def it_gives_access_to_followed_hyperlink_color(self, color_scheme):
        assert isinstance(color_scheme.folHlink, ColorFormat)


    # fixtures ---------------------------------------------
    @pytest.fixture(
        params=[
            ("a:clrScheme{name=Foobar}", "Foobar")
        ]
    )
    def name_fixture(self, request):
        clr_scheme_cxml, expected_value = request.param
        color_scheme = ColorScheme(element(clr_scheme_cxml))
        return color_scheme, expected_value

    # fixture components ---------------------------------------------
    
    @pytest.fixture
    def color_scheme(self):
        return ColorScheme(element("a:clrScheme/(a:dk1,a:lt1,a:dk2,a:lt2,a:accent1,a:accent2,a:accent3,a:accent4,a:accent5,a:accent6,a:hlink,a:folHlink)"))

class DescribeFontScheme:
    def it_gives_access_to_its_major_font(self, font_scheme):
        assert isinstance(font_scheme.major_font, FontCollection)

    def it_gives_access_to_its_minor_font(self, font_scheme):
        assert isinstance(font_scheme.minor_font, FontCollection)

    # fixtures ---------------------------------------------

    # fixture components ---------------------------------------------

    @pytest.fixture
    def font_scheme(self):
        return FontScheme(element("a:fontScheme/(a:majorFont,a:minorFont)"))

class DescribeFormatScheme:
    def it_knows_its_name(self, name_fixture):
        format_scheme, expected_value = name_fixture
        assert format_scheme.name == expected_value

    def it_knows_its_fill_style_list(self, format_scheme):
        fill_list = format_scheme.fill_style_list
        assert isinstance(fill_list, tuple)
        for fill in fill_list:
            assert isinstance(fill, FillFormat)
        

    def it_knows_its_line_style_list(self, format_scheme):
        line_style_list = format_scheme.line_style_list
        assert isinstance(line_style_list, tuple)
        for line in line_style_list:
            assert isinstance(line, LineStyle)
        

    def it_knows_its_effect_style_list(self, format_scheme):
        effect_style_list = format_scheme.effect_style_list
        assert isinstance(effect_style_list, list)
        assert effect_style_list == []

    def it_knows_its_background_fill_style_list(self, format_scheme):
        bg_fill_list = format_scheme.background_fill_style_list
        assert isinstance(bg_fill_list, tuple)
        for fill in bg_fill_list:
            assert isinstance(fill, FillFormat)
        

    # fixtures ---------------------------------------------

    @pytest.fixture(
        params=[
            ("a:fmtScheme", ""),
            ("a:fmtScheme{name=Foobar}", "Foobar")
        ]
    )
    def name_fixture(self, request):
        format_scheme_cxml, expected_value = request.param
        format_scheme = FormatScheme(element(format_scheme_cxml))
        return format_scheme, expected_value




    # fixture components ---------------------------------------------
    @pytest.fixture
    def format_scheme(self, request):
        return FormatScheme(element("a:fmtScheme/(a:fillStyleLst/(a:noFill,a:solidFill,a:solidFill),a:lnStyleLst/(a:ln,a:ln,a:ln),a:bgFillStyleLst/(a:noFill,a:solidFill,a:solidFill))"))

class DescribeFontCollection:
    def it_provides_access_to_latin_font(self, font_collection):
        assert isinstance(font_collection.latin, TextFont)

    def it_provides_access_to_ea_font(self, font_collection):
        assert isinstance(font_collection.east_asian, TextFont)

    def it_provides_access_to_cs_font(self, font_collection):
        assert isinstance(font_collection.complex_script, TextFont)


    # fixtures ---------------------------------------------

    # fixture components ---------------------------------------------

    @pytest.fixture
    def font_collection(self):
        return FontCollection(element("a:majorFont/(a:latin,a:ea,a:cs)"))
