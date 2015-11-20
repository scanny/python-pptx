# encoding: utf-8

"""
Test suite for pptx.oxml.dml module.
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.dml.color import CT_Percentage, CT_SchemeColor, CT_SRgbColor
from pptx.oxml.ns import qn

from .unitdata.dml import a_lumMod, a_lumOff, a_schemeClr, an_srgbClr


class Describe_BaseColorElement(object):

    def it_can_get_the_lumMod_child_element_if_there_is_one(
            self, schemeClr, schemeClr_with_lumMod, lumMod):
        assert schemeClr.lumMod is None
        assert schemeClr_with_lumMod.lumMod is lumMod

    def it_can_get_the_lumOff_child_element_if_there_is_one(
            self, schemeClr, schemeClr_with_lumOff, lumOff):
        assert schemeClr.lumOff is None
        assert schemeClr_with_lumOff.lumOff is lumOff

    def it_can_remove_existing_lumMod_and_lumOff_child_elements(
            self, schemeClr_with_lumMod, schemeClr_with_lumOff,
            schemeClr_xml):
        schemeClr_with_lumMod.clear_lum()
        schemeClr_with_lumOff.clear_lum()
        assert schemeClr_with_lumMod.xml == schemeClr_xml
        assert schemeClr_with_lumOff.xml == schemeClr_xml

    def it_can_add_a_lumMod_child_element(
            self, schemeClr, schemeClr_with_lumMod_xml):
        lumMod = schemeClr.add_lumMod(0.75)
        assert schemeClr.xml == schemeClr_with_lumMod_xml
        assert schemeClr.find(qn('a:lumMod')) == lumMod

    def it_can_add_a_lumOff_child_element(
            self, schemeClr, schemeClr_with_lumOff_xml):
        lumOff = schemeClr.add_lumOff(0.4)
        assert schemeClr.xml == schemeClr_with_lumOff_xml
        assert schemeClr.find(qn('a:lumOff')) == lumOff

    # fixtures ---------------------------------------------

    @pytest.fixture
    def lumMod(self):
        return a_lumMod().with_nsdecls().element

    @pytest.fixture
    def lumOff(self):
        return a_lumOff().with_nsdecls().element

    @pytest.fixture
    def schemeClr(self):
        return a_schemeClr().with_nsdecls().element

    @pytest.fixture
    def schemeClr_xml(self):
        return a_schemeClr().with_nsdecls().xml()

    @pytest.fixture
    def schemeClr_with_lumMod(self, lumMod):
        schemeClr = a_schemeClr().with_nsdecls().element
        schemeClr.append(lumMod)
        return schemeClr

    @pytest.fixture
    def schemeClr_with_lumMod_xml(self):
        lumMod_bldr = a_lumMod().with_val(75000)
        return a_schemeClr().with_nsdecls().with_child(lumMod_bldr).xml()

    @pytest.fixture
    def schemeClr_with_lumOff_xml(self):
        lumOff_bldr = a_lumOff().with_val(40000)
        return a_schemeClr().with_nsdecls().with_child(lumOff_bldr).xml()

    @pytest.fixture
    def schemeClr_with_lumOff(self, lumOff):
        schemeClr = a_schemeClr().with_nsdecls().element
        schemeClr.append(lumOff)
        return schemeClr


class DescribeCT_Percentage(object):

    def it_is_used_by_the_parser_for_a_lumOff_element(self, lumOff):
        assert isinstance(lumOff, CT_Percentage)

    def it_is_used_by_the_parser_for_a_lumMod_element(self, lumMod):
        assert isinstance(lumMod, CT_Percentage)

    def it_knows_the_percentage_value(self, ct_percentage):
        assert ct_percentage.val == 0.99999

    # fixtures ---------------------------------------------

    @pytest.fixture
    def ct_percentage(self):
        return a_lumMod().with_nsdecls().with_val('99999').element

    @pytest.fixture
    def lumMod(self):
        return a_lumMod().with_nsdecls().with_val('33333').element

    @pytest.fixture
    def lumOff(self):
        return a_lumOff().with_nsdecls().with_val('66666').element


class DescribeCT_SchemeColor(object):

    def it_is_used_by_the_parser_for_a_schemeClr_element(self, schemeClr):
        assert isinstance(schemeClr, CT_SchemeColor)

    def it_knows_the_theme_color_idx(self, schemeClr):
        assert schemeClr.val == MSO_THEME_COLOR.BACKGROUND_1

    def it_can_set_the_scheme_color_value(self, schemeClr, schemeClr_xml):
        schemeClr.val = MSO_THEME_COLOR.ACCENT_1
        assert schemeClr.xml == schemeClr_xml

    # fixtures ---------------------------------------------

    @pytest.fixture
    def schemeClr(self):
        return a_schemeClr().with_nsdecls().with_val('bg1').element

    @pytest.fixture
    def schemeClr_xml(self):
        return a_schemeClr().with_nsdecls().with_val('accent1').xml()


class DescribeCT_SRgbColor(object):

    def it_is_used_by_the_parser_for_an_srgbClr_element(self, srgbClr):
        assert isinstance(srgbClr, CT_SRgbColor)

    def it_knows_the_rgb_str_value(self, srgbClr):
        assert srgbClr.val == '123456'

    def it_can_set_the_rgb_str_value(self, srgbClr, srgbClr_xml):
        srgbClr.val = '987654'
        assert srgbClr.xml == srgbClr_xml

    # fixtures ---------------------------------------------

    @pytest.fixture
    def srgbClr(self):
        return an_srgbClr().with_nsdecls().with_val('123456').element

    @pytest.fixture
    def srgbClr_xml(self):
        return an_srgbClr().with_nsdecls().with_val('987654').xml()
