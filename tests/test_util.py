# encoding: utf-8

"""Test suite for pptx.util module."""

from __future__ import absolute_import

import platform
import pytest

from pptx.util import (
    BaseLength, Centipoints, Cm, Collection, Emu, Inches, Mm, Px, to_unicode
)

from .unitutil import TestCase


def test_to_unicode_raises_on_non_string():
    """to_unicode(text) raises on *text* not a string"""
    with pytest.raises(TypeError):
        to_unicode(999)


class TestCollection(TestCase):
    """Test Collection"""
    def setUp(self):
        self.collection = Collection()

    def test_indexable(self):
        """Collection is indexable (e.g. no TypeError on 'collection[0]')"""
        # verify ----------------------
        try:
            self.collection[0]
        except TypeError:
            msg = "'Collection' object does not support indexing"
            self.fail(msg)
        except IndexError:
            pass

    def test_is_container(self):
        """Collection is container (e.g. 'x in collection' works)"""
        # verify ----------------------
        try:
            1 in self.collection
        except TypeError:
            msg = "'Collection' object is not container"
            self.fail(msg)

    def test_iterable(self):
        """Collection is iterable"""
        # verify ----------------------
        try:
            for x in self.collection:
                pass
        except TypeError:
            msg = "'Collection' object is not iterable"
            self.fail(msg)

    def test_sized(self):
        """Collection is sized (e.g. 'len(collection)' works)"""
        # verify ----------------------
        try:
            len(self.collection)
        except TypeError:
            msg = "object of type 'Collection' has no len()"
            self.fail(msg)

    def test__values_property_empty_on_construction(self):
        """Collection._values property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.collection, '_values', 0)


class DescribeLength(object):

    def it_can_construct_from_convenient_units(self, construct_fixture):
        UnitCls, units_val, emu = construct_fixture
        length = UnitCls(units_val)
        assert isinstance(length, BaseLength)
        assert length == emu

    def it_can_self_convert_to_convenient_units(self, units_fixture):
        emu, units_prop_name, expected_length_in_units = units_fixture
        length = BaseLength(emu)
        length_in_units = getattr(length, units_prop_name)
        assert length_in_units == expected_length_in_units

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (BaseLength,  914400,  914400),
        (Inches,      1.1,    1005840),
        (Centipoints, 12.5,      1587),
        (Cm,          2.53,    910799),
        (Emu,         9144.9,    9144),
        (Mm,          13.8,    496800),
        (Px,          10,
         95250 if platform.system() == 'Windows' else 127000),
    ])
    def construct_fixture(self, request):
        UnitCls, units_val, emu = request.param
        return UnitCls, units_val, emu

    @pytest.fixture(params=[
        (914400, 'inches', 1.0),
        (914400, 'centipoints', 7200.0),
        (914400, 'cm', 2.54),
        (914400, 'emu', 914400),
        (914400, 'mm', 25.4),
        (914400, 'px', 96 if platform.system() == 'Windows' else 72),
    ])
    def units_fixture(self, request):
        emu, units_prop_name, expected_length_in_units = request.param
        return emu, units_prop_name, expected_length_in_units
