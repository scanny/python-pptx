# encoding: utf-8

"""
Test suite for pptx.util module.
"""

from __future__ import absolute_import

import pytest

from pptx.compat import to_unicode
from pptx.util import Length, Centipoints, Cm, Emu, Inches, Mm, Pt


def test_to_unicode_raises_on_non_string():
    """to_unicode(text) raises on *text* not a string"""
    with pytest.raises(TypeError):
        to_unicode(999)


class DescribeLength(object):
    def it_can_construct_from_convenient_units(self, construct_fixture):
        UnitCls, units_val, emu = construct_fixture
        length = UnitCls(units_val)
        assert isinstance(length, Length)
        assert length == emu

    def it_can_self_convert_to_convenient_units(self, units_fixture):
        emu, units_prop_name, expected_length_in_units = units_fixture
        length = Length(emu)
        length_in_units = getattr(length, units_prop_name)
        assert length_in_units == expected_length_in_units

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            (Length, 914400, 914400),
            (Inches, 1.1, 1005840),
            (Centipoints, 12.5, 1587),
            (Cm, 2.53, 910799),
            (Emu, 9144.9, 9144),
            (Mm, 13.8, 496800),
            (Pt, 24.5, 311150),
        ]
    )
    def construct_fixture(self, request):
        UnitCls, units_val, emu = request.param
        return UnitCls, units_val, emu

    @pytest.fixture(
        params=[
            (914400, "inches", 1.0),
            (914400, "centipoints", 7200.0),
            (914400, "cm", 2.54),
            (914400, "emu", 914400),
            (914400, "mm", 25.4),
            (914400, "pt", 72.0),
        ]
    )
    def units_fixture(self, request):
        emu, units_prop_name, expected_length_in_units = request.param
        return emu, units_prop_name, expected_length_in_units
