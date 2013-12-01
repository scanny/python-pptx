# encoding: utf-8

"""Test suite for pptx.text module."""

from __future__ import absolute_import

import pytest

from pptx.dml.core import FillFormat, RGBColor

from ..oxml.unitdata.dml import a_gradFill, a_solidFill, an_spPr
from ..unitutil import actual_xml


class DescribeFillFormat(object):

    def it_can_set_the_fill_type_to_solid(self, set_solid_fixture_):
        fill, spPr_with_solidFill_xml = set_solid_fixture_
        fill.solid()
        assert actual_xml(fill._xPr) == spPr_with_solidFill_xml

    # fixtures -------------------------------------------------------

    def _fill_type_cases():
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

    @pytest.fixture(params=_fill_type_cases())
    def set_solid_fixture_(self, request, spPr_with_solidFill_xml):
        spPr = request.param
        return FillFormat(spPr), spPr_with_solidFill_xml

    @pytest.fixture
    def spPr_with_solidFill_xml(self):
        return (
            an_spPr().with_nsdecls()
                     .with_child(a_solidFill())
                     .xml()
        )


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
