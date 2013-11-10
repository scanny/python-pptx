# encoding: utf-8

"""Test suite for pptx.text module."""

from __future__ import absolute_import

import pytest

from pptx.dml.core import RGBColor


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
