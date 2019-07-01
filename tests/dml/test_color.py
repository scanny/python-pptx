# encoding: utf-8

"""
Test suite for pptx.text module.
"""

from __future__ import absolute_import

import pytest

from pptx.dml.color import ColorFormat, RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR

from ..oxml.unitdata.dml import (
    a_lumMod,
    a_lumOff,
    a_prstClr,
    a_schemeClr,
    a_solidFill,
    a_sysClr,
    an_hslClr,
    an_scrgbClr,
    an_srgbClr,
)


class DescribeColorFormat(object):
    def it_knows_the_type_of_its_color(self, color_type_fixture_):
        color_format, color_type = color_type_fixture_
        assert color_format.type == color_type

    def it_knows_the_RGB_value_of_an_RGB_color(self, rgb_color_format):
        color_format = rgb_color_format
        assert color_format.rgb == RGBColor(0x12, 0x34, 0x56)

    def it_raises_on_rgb_get_for_colors_other_than_rgb(self, rgb_raise_fixture_):
        color_format, exception_type = rgb_raise_fixture_
        with pytest.raises(exception_type):
            color_format.rgb

    def it_knows_the_theme_color_of_a_theme_color(self, get_theme_color_fixture_):
        color_format, theme_color = get_theme_color_fixture_
        assert color_format.theme_color == theme_color

    def it_raises_on_theme_color_get_for_NoneColor(self, _NoneColor_color_format):
        with pytest.raises(AttributeError):
            _NoneColor_color_format.theme_color

    def it_knows_its_brightness_adjustment(self, color_format_with_brightness):
        color_format, expected_brightness = color_format_with_brightness
        assert color_format.brightness == expected_brightness

    def it_can_set_itself_to_an_RGB_color(self, set_rgb_fixture_):
        color_format, rgb_color, expected_xml = set_rgb_fixture_
        color_format.rgb = rgb_color
        assert color_format._xFill.xml == expected_xml

    def it_raises_on_assign_non_RGBColor_type_to_rgb(self, rgb_color_format):
        color_format = rgb_color_format
        with pytest.raises(ValueError):
            color_format.rgb = (0x12, 0x34, 0x56)

    def it_can_set_itself_to_a_theme_color(self, set_theme_color_fixture_):
        color_format, theme_color, expected_xml = set_theme_color_fixture_
        color_format.theme_color = theme_color
        assert color_format._xFill.xml == expected_xml

    def it_can_set_its_brightness_adjustment(self, set_brightness_fixture_):
        color_format, brightness, expected_xml = set_brightness_fixture_
        color_format.brightness = brightness
        assert color_format._xFill.xml == expected_xml

    def it_raises_on_attempt_to_set_brightness_out_of_range(self, rgb_color_format):
        with pytest.raises(ValueError):
            rgb_color_format.brightness = 1.1
        with pytest.raises(ValueError):
            rgb_color_format.brightness = -1.1

    def it_raises_on_attempt_to_set_brightness_on_None_color_type(
        self, color_format_having_none_color_type
    ):
        color_format = color_format_having_none_color_type
        with pytest.raises(ValueError):
            color_format.brightness = 0.5

    # fixtures ---------------------------------------------

    @pytest.fixture
    def color_format_having_none_color_type(self):
        solidFill = a_solidFill().with_nsdecls().element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format

    @pytest.fixture(params=["hsl", "prst", "scheme", "scrgb", "srgb", "sys"])
    def color_format_with_brightness(self, request):
        mapping = {
            "hsl": (an_hslClr, 55000, 45000, 0.45),
            "prst": (a_prstClr, None, None, 0.0),
            "scheme": (a_schemeClr, 15000, None, -0.85),
            "scrgb": (an_scrgbClr, 15000, 85000, 0.85),
            "srgb": (an_srgbClr, None, None, 0.0),
            "sys": (a_sysClr, 23000, None, -0.77),
        }
        xClr_bldr_fn, lumMod, lumOff, exp_brightness = mapping[request.param]
        xClr_bldr = xClr_bldr_fn()
        if lumMod is not None:
            xClr_bldr.with_child(a_lumMod().with_val(lumMod))
        if lumOff is not None:
            xClr_bldr.with_child(a_lumOff().with_val(lumOff))
        solidFill = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format, exp_brightness

    @pytest.fixture(params=["none", "hsl", "prst", "scheme", "scrgb", "srgb", "sys"])
    def color_type_fixture_(self, request):
        mapping = {
            "none": (None, None),
            "hsl": (an_hslClr, MSO_COLOR_TYPE.HSL),
            "prst": (a_prstClr, MSO_COLOR_TYPE.PRESET),
            "srgb": (an_srgbClr, MSO_COLOR_TYPE.RGB),
            "scheme": (a_schemeClr, MSO_COLOR_TYPE.SCHEME),
            "scrgb": (an_scrgbClr, MSO_COLOR_TYPE.SCRGB),
            "sys": (a_sysClr, MSO_COLOR_TYPE.SYSTEM),
        }
        xClr_bldr_fn, color_type = mapping[request.param]
        solidFill_bldr = a_solidFill().with_nsdecls()
        if xClr_bldr_fn is not None:
            solidFill_bldr.with_child(xClr_bldr_fn())
        solidFill = solidFill_bldr.element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format, color_type

    @pytest.fixture
    def _NoneColor_color_format(self):
        solidFill = a_solidFill().with_nsdecls().element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format

    @pytest.fixture
    def rgb_color_format(self):
        srgbClr_bldr = an_srgbClr().with_val("123456")
        solidFill = a_solidFill().with_nsdecls().with_child(srgbClr_bldr).element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format

    @pytest.fixture(params=["none", "hsl", "prst", "scheme", "scrgb", "sys"])
    def rgb_raise_fixture_(self, request):
        mapping = {
            "none": (None, AttributeError),
            "hsl": (an_hslClr, AttributeError),
            "prst": (a_prstClr, AttributeError),
            "scheme": (a_schemeClr, AttributeError),
            "scrgb": (an_scrgbClr, AttributeError),
            "sys": (a_sysClr, AttributeError),
        }
        xClr_bldr_fn, exception_type = mapping[request.param]
        solidFill_bldr = a_solidFill().with_nsdecls()
        if xClr_bldr_fn is not None:
            solidFill_bldr.with_child(xClr_bldr_fn())
        solidFill = solidFill_bldr.element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format, exception_type

    @pytest.fixture(
        params=[
            "0 to 0",
            "0 to -0.4",
            "0.15 to 0.25",
            "0.15 to -0.15",
            "-0.25 to 0.4",
            "-0.3 to -0.4",
            "-0.4 to 0",
        ]
    )
    def set_brightness_fixture_(self, request):
        mapping = {
            "0 to 0": (an_srgbClr, None, None, 0, None, None),
            "0 to -0.4": (an_hslClr, None, None, -0.4, 60000, None),
            "0.15 to 0.25": (a_prstClr, 85000, 15000, 0.25, 75000, 25000),
            "0.15 to -0.15": (a_schemeClr, 85000, 15000, -0.15, 85000, None),
            "-0.25 to 0.4": (an_scrgbClr, 75000, None, 0.4, 60000, 40000),
            "-0.3 to -0.4": (an_srgbClr, 70000, None, -0.4, 60000, None),
            "-0.4 to 0": (a_sysClr, 60000, None, 0, None, None),
        }
        xClr_bldr_fn, mod_in, off_in, brightness, mod_out, off_out = mapping[
            request.param
        ]

        xClr_bldr = xClr_bldr_fn()
        if mod_in is not None:
            xClr_bldr.with_child(a_lumMod().with_val(mod_in))
        if off_in is not None:
            xClr_bldr.with_child(a_lumOff().with_val(off_in))
        solidFill = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)

        xClr_bldr = xClr_bldr_fn()
        if mod_out is not None:
            xClr_bldr.with_child(a_lumMod().with_val(mod_out))
        if off_out is not None:
            xClr_bldr.with_child(a_lumOff().with_val(off_out))
        expected_xml = a_solidFill().with_nsdecls().with_child(xClr_bldr).xml()

        return color_format, brightness, expected_xml

    @pytest.fixture(params=["none", "hsl", "prst", "scheme", "scrgb", "srgb", "sys"])
    def set_rgb_fixture_(self, request):
        mapping = {
            "none": None,
            "hsl": an_hslClr,
            "prst": a_prstClr,
            "scheme": a_schemeClr,
            "scrgb": an_scrgbClr,
            "srgb": an_srgbClr,
            "sys": a_sysClr,
        }
        xClr_bldr_fn = mapping[request.param]
        solidFill_bldr = a_solidFill().with_nsdecls()
        if xClr_bldr_fn is not None:
            solidFill_bldr.with_child(xClr_bldr_fn())
        solidFill = solidFill_bldr.element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        rgb_color = RGBColor(0x12, 0x34, 0x56)
        expected_xml = (
            a_solidFill()
            .with_nsdecls()
            .with_child(an_srgbClr().with_val("123456"))
            .xml()
        )
        return color_format, rgb_color, expected_xml

    @pytest.fixture(params=["none", "hsl", "prst", "scheme", "scrgb", "srgb", "sys"])
    def set_theme_color_fixture_(self, request):
        mapping = {
            "none": None,
            "hsl": an_hslClr,
            "prst": a_prstClr,
            "scheme": a_schemeClr,
            "scrgb": an_scrgbClr,
            "srgb": an_srgbClr,
            "sys": a_sysClr,
        }
        xClr_bldr_fn = mapping[request.param]
        solidFill_bldr = a_solidFill().with_nsdecls()
        if xClr_bldr_fn is not None:
            solidFill_bldr.with_child(xClr_bldr_fn())
        solidFill = solidFill_bldr.element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        theme_color = MSO_THEME_COLOR.ACCENT_6
        expected_xml = (
            a_solidFill()
            .with_nsdecls()
            .with_child(a_schemeClr().with_val("accent6"))
            .xml()
        )
        return color_format, theme_color, expected_xml

    @pytest.fixture(params=["hsl", "prst", "scheme", "scrgb", "srgb", "sys"])
    def get_theme_color_fixture_(self, request):
        mapping = {
            "hsl": (an_hslClr, MSO_THEME_COLOR.NOT_THEME_COLOR),
            "prst": (a_prstClr, MSO_THEME_COLOR.NOT_THEME_COLOR),
            "scheme": (a_schemeClr, MSO_THEME_COLOR.ACCENT_1),
            "scrgb": (an_scrgbClr, MSO_THEME_COLOR.NOT_THEME_COLOR),
            "srgb": (an_srgbClr, MSO_THEME_COLOR.NOT_THEME_COLOR),
            "sys": (a_sysClr, MSO_THEME_COLOR.NOT_THEME_COLOR),
        }
        xClr_bldr_fn, theme_color = mapping[request.param]
        xClr_bldr = xClr_bldr_fn()
        if theme_color != MSO_THEME_COLOR.NOT_THEME_COLOR:
            xClr_bldr.with_val("accent1")
        solidFill = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format, theme_color


class DescribeRGBColor(object):
    def it_is_natively_constructed_using_three_ints_0_to_255(self):
        RGBColor(0x12, 0x34, 0x56)
        with pytest.raises(ValueError):
            RGBColor("12", "34", "56")
        with pytest.raises(ValueError):
            RGBColor(-1, 34, 56)
        with pytest.raises(ValueError):
            RGBColor(12, 256, 56)

    def it_can_construct_from_a_hex_string_rgb_value(self):
        rgb = RGBColor.from_string("123456")
        assert rgb == RGBColor(0x12, 0x34, 0x56)

    def it_can_provide_a_hex_string_rgb_value(self):
        assert str(RGBColor(0x12, 0x34, 0x56)) == "123456"
