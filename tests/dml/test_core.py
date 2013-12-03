# encoding: utf-8

"""
Test suite for pptx.text module.
"""

from __future__ import absolute_import

import pytest

from pptx.dml.core import ColorFormat, FillFormat, RGBColor
from pptx.enum import (
    MSO_COLOR_TYPE, MSO_FILL_TYPE as MSO_FILL, MSO_THEME_COLOR
)

from ..oxml.unitdata.dml import (
    a_blipFill, a_gradFill, a_grpFill, a_lumMod, a_lumOff, a_noFill,
    a_pattFill, a_prstClr, a_schemeClr, a_solidFill, a_sysClr, an_hslClr,
    an_spPr, an_scrgbClr, an_srgbClr
)
from ..unitutil import actual_xml


class DescribeColorFormat(object):

    def it_knows_the_type_of_its_color(self, color_type_fixture_):
        color_format, color_type = color_type_fixture_
        print(actual_xml(color_format._xFill))
        assert color_format.type == color_type

    def it_knows_the_RGB_value_of_an_RGB_color(self, rgb_color_format):
        color_format = rgb_color_format
        assert color_format.rgb == RGBColor(0x12, 0x34, 0x56)

    def it_raises_on_rgb_get_for_colors_other_than_rgb(
            self, rgb_raise_fixture_):
        color_format, exception_type = rgb_raise_fixture_
        with pytest.raises(exception_type):
            color_format.rgb

    def it_knows_the_theme_color_of_a_theme_color(self, theme_color_format):
        color_format = theme_color_format
        assert color_format.theme_color == MSO_THEME_COLOR.ACCENT_1

    def it_raises_on_theme_color_get_for_colors_other_than_schemeClr(
            self, theme_color_raise_fixture_):
        color_format, exception_type = theme_color_raise_fixture_
        with pytest.raises(exception_type):
            color_format.theme_color

    def it_knows_its_brightness_adjustment(
            self, color_format_with_brightness):
        color_format, expected_brightness = color_format_with_brightness
        print(actual_xml(color_format._xFill))
        assert color_format.brightness == expected_brightness

    def it_can_set_itself_to_an_RGB_color(self, set_rgb_fixture_):
        color_format, rgb_color, expected_xml = set_rgb_fixture_
        color_format.rgb = rgb_color
        assert actual_xml(color_format._xFill) == expected_xml

    def it_raises_on_assign_non_RGBColor_type_to_rgb(self, rgb_color_format):
        color_format = rgb_color_format
        with pytest.raises(ValueError):
            color_format.rgb = (0x12, 0x34, 0x56)

    # def it_can_set_itself_to_a_theme_color(self):
    # def it_can_set_its_brightness_adjustment(self):
    # def it_raises_on_attempt_to_set_brightness_out_of_range(self):
    # def it_raises_on_attempt_to_set_brightness_on_None_color_type(self):

    # fixtures ---------------------------------------------

    @pytest.fixture(params=['hsl', 'prst', 'scheme', 'scrgb', 'srgb', 'sys'])
    def color_format_with_brightness(self, request):
        mapping = {
            'hsl':    (an_hslClr,   55000, 45000,  0.45),
            'prst':   (a_prstClr,   None,  None,   0.0),
            'scheme': (a_schemeClr, 15000, None,  -0.85),
            'scrgb':  (an_scrgbClr, 15000, 85000,  0.85),
            'srgb':   (an_srgbClr,  None,  None,   0.0),
            'sys':    (a_sysClr,    23000, None,  -0.77),
        }
        xClr_bldr_fn, lumMod, lumOff, exp_brightness = mapping[request.param]
        xClr_bldr = xClr_bldr_fn()
        if lumMod is not None:
            xClr_bldr.with_child(a_lumMod().with_val(lumMod))
        if lumOff is not None:
            xClr_bldr.with_child(a_lumOff().with_val(lumOff))
        solidFill = (
            a_solidFill().with_nsdecls()
                         .with_child(xClr_bldr)
                         .element
        )
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format, exp_brightness

    @pytest.fixture(params=[
        'none', 'hsl', 'prst', 'scheme', 'scrgb', 'srgb', 'sys'
    ])
    def color_type_fixture_(self, request):
        mapping = {
            'none':   (None,        None),
            'hsl':    (an_hslClr,   MSO_COLOR_TYPE.HSL),
            'prst':   (a_prstClr,   MSO_COLOR_TYPE.PRESET),
            'srgb':   (an_srgbClr,  MSO_COLOR_TYPE.RGB),
            'scheme': (a_schemeClr, MSO_COLOR_TYPE.SCHEME),
            'scrgb':  (an_scrgbClr, MSO_COLOR_TYPE.SCRGB),
            'sys':    (a_sysClr,    MSO_COLOR_TYPE.SYSTEM),
        }
        xClr_bldr_fn, color_type = mapping[request.param]
        solidFill_bldr = a_solidFill().with_nsdecls()
        if xClr_bldr_fn is not None:
            solidFill_bldr.with_child(xClr_bldr_fn())
        solidFill = solidFill_bldr.element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format, color_type

    @pytest.fixture
    def rgb_color_format(self):
        srgbClr_bldr = an_srgbClr().with_val('123456')
        solidFill = (
            a_solidFill().with_nsdecls().with_child(srgbClr_bldr).element
        )
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format

    @pytest.fixture(params=['none', 'hsl', 'prst', 'scheme', 'scrgb', 'sys'])
    def rgb_raise_fixture_(self, request):
        mapping = {
            'none':   (None,        AttributeError),
            'hsl':    (an_hslClr,   AttributeError),
            'prst':   (a_prstClr,   AttributeError),
            'scheme': (a_schemeClr, AttributeError),
            'scrgb':  (an_scrgbClr, AttributeError),
            'sys':    (a_sysClr,    AttributeError),
        }
        xClr_bldr_fn, exception_type = mapping[request.param]
        solidFill_bldr = a_solidFill().with_nsdecls()
        if xClr_bldr_fn is not None:
            solidFill_bldr.with_child(xClr_bldr_fn())
        solidFill = solidFill_bldr.element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format, exception_type

    @pytest.fixture(params=[
        'none', 'hsl', 'prst', 'scheme', 'scrgb', 'srgb', 'sys'
    ])
    def set_rgb_fixture_(self, request):
        mapping = {
            'none':   None,
            'hsl':    an_hslClr,
            'prst':   a_prstClr,
            'scheme': a_schemeClr,
            'scrgb':  an_scrgbClr,
            'srgb':   an_srgbClr,
            'sys':    a_sysClr,
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
            .with_child(an_srgbClr().with_val('123456'))
            .xml()
        )
        return color_format, rgb_color, expected_xml

    @pytest.fixture
    def theme_color_format(self):
        schemeClr_bldr = a_schemeClr().with_val('accent1')
        solidFill = (
            a_solidFill().with_nsdecls().with_child(schemeClr_bldr).element
        )
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format

    @pytest.fixture(params=['none', 'hsl', 'prst', 'scrgb', 'srgb', 'sys'])
    def theme_color_raise_fixture_(self, request):
        mapping = {
            'none':  (None,        AttributeError),
            'hsl':   (an_hslClr,   AttributeError),
            'prst':  (a_prstClr,   AttributeError),
            'scrgb': (an_scrgbClr, AttributeError),
            'srgb':  (an_srgbClr,  AttributeError),
            'sys':   (a_sysClr,    AttributeError),
        }
        xClr_bldr_fn, exception_type = mapping[request.param]
        solidFill_bldr = a_solidFill().with_nsdecls()
        if xClr_bldr_fn is not None:
            solidFill_bldr.with_child(xClr_bldr_fn())
        solidFill = solidFill_bldr.element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format, exception_type


class DescribeFillFormat(object):

    def it_can_set_the_fill_type_to_solid(self, set_solid_fixture_):
        fill, spPr_with_solidFill_xml = set_solid_fixture_
        fill.solid()
        assert actual_xml(fill._xPr) == spPr_with_solidFill_xml

    def it_knows_the_type_of_fill_it_is(self, fill_type_fixture_):
        fill_format, fill_type = fill_type_fixture_
        print(actual_xml(fill_format._xPr))
        assert fill_format.fill_type == fill_type

    def it_provides_access_to_the_foreground_color_object(
            self, fore_color_fixture_):
        fill_format, fore_color_type = fore_color_fixture_
        print(actual_xml(fill_format._xPr))
        assert isinstance(fill_format.fore_color, fore_color_type)

    def it_raises_on_fore_color_get_for_fill_types_that_dont_have_one(
            self, fore_color_raise_fixture_):
        fill_format, exception_type = fore_color_raise_fixture_
        with pytest.raises(exception_type):
            fill_format.fore_color

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        'none', 'no', 'solid', 'grad', 'blip', 'patt', 'grp'
    ])
    def fill_type_fixture_(self, request):
        mapping = {
            'none':  ('_spPr_bldr',                None),
            'grad':  ('_spPr_with_gradFill_bldr',  MSO_FILL.GRADIENT),
            'solid': ('_spPr_with_solidFill_bldr', MSO_FILL.SOLID),
            'no':    ('_spPr_with_noFill_bldr',    MSO_FILL.BACKGROUND),
            'blip':  ('_spPr_with_blipFill_bldr',  MSO_FILL.PICTURE),
            'patt':  ('_spPr_with_pattFill_bldr',  MSO_FILL.PATTERNED),
            'grp':   ('_spPr_with_grpFill_bldr',   MSO_FILL.GROUP),
        }
        spPr_bldr_name, fill_type = mapping[request.param]
        spPr_bldr = request.getfuncargvalue(spPr_bldr_name)
        spPr = spPr_bldr.element
        fill_format = FillFormat.from_fill_parent(spPr)
        return fill_format, fill_type

    @pytest.fixture(params=['solid'])
    def fore_color_fixture_(self, request):
        mapping = {
            'solid': ('_spPr_with_solidFill_bldr', ColorFormat),
        }
        spPr_bldr_name, fore_color_type = mapping[request.param]
        spPr_bldr = request.getfuncargvalue(spPr_bldr_name)
        spPr = spPr_bldr.element
        fill_format = FillFormat.from_fill_parent(spPr)
        return fill_format, fore_color_type

    @pytest.fixture(params=['none', 'blip', 'grad', 'grp', 'patt'])
    def fore_color_raise_fixture_(self, request):
        mapping = {
            'none':  ('_spPr_bldr',                TypeError),
            'blip':  ('_spPr_with_blipFill_bldr',  TypeError),
            'grad':  ('_spPr_with_gradFill_bldr',  NotImplementedError),
            'grp':   ('_spPr_with_grpFill_bldr',   TypeError),
            'patt':  ('_spPr_with_pattFill_bldr',  NotImplementedError),
        }
        spPr_bldr_name, exception_type = mapping[request.param]
        spPr_bldr = request.getfuncargvalue(spPr_bldr_name)
        spPr = spPr_bldr.element
        fill_format = FillFormat.from_fill_parent(spPr)
        return fill_format, exception_type

    def _solid_fill_cases():
        # no fill type yet
        spPr = an_spPr().with_nsdecls().element
        # non-solid fill type present
        spPr_with_gradFill = (
            an_spPr().with_nsdecls()
                     .with_child(a_gradFill())
                     .element
        )
        # solidFill already present
        spPr_with_solidFill = (
            an_spPr().with_nsdecls()
                     .with_child(a_solidFill())
                     .element
        )
        return [spPr, spPr_with_gradFill, spPr_with_solidFill]

    @pytest.fixture(params=_solid_fill_cases())
    def set_solid_fixture_(self, request, spPr_with_solidFill_xml):
        spPr = request.param
        fill = FillFormat.from_fill_parent(spPr)
        return fill, spPr_with_solidFill_xml

    @pytest.fixture
    def spPr_with_solidFill_xml(self):
        return (
            an_spPr().with_nsdecls()
                     .with_child(a_solidFill())
                     .xml()
        )

    @pytest.fixture
    def _spPr_bldr(self, request):
        return an_spPr().with_nsdecls()

    @pytest.fixture
    def _spPr_with_noFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_noFill())

    @pytest.fixture
    def _spPr_with_solidFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_solidFill())

    @pytest.fixture
    def _spPr_with_gradFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_gradFill())

    @pytest.fixture
    def _spPr_with_blipFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_blipFill())

    @pytest.fixture
    def _spPr_with_pattFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_pattFill())

    @pytest.fixture
    def _spPr_with_grpFill_bldr(self, request):
        return an_spPr().with_nsdecls().with_child(a_grpFill())


class DescribeRGBColor(object):

    def it_is_natively_constructed_using_three_ints_0_to_255(self):
        RGBColor(0x12, 0x34, 0x56)
        with pytest.raises(ValueError):
            RGBColor('12', '34', '56')
        with pytest.raises(ValueError):
            RGBColor(-1, 34, 56)
        with pytest.raises(ValueError):
            RGBColor(12, 256, 56)

    def it_can_construct_from_a_hex_string_rgb_value(self):
        rgb = RGBColor.from_string('123456')
        assert rgb == RGBColor(0x12, 0x34, 0x56)

    def it_can_provide_a_hex_string_rgb_value(self):
        assert str(RGBColor(0x12, 0x34, 0x56)) == '123456'
