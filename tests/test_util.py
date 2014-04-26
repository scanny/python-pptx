# encoding: utf-8

"""Test suite for pptx.util module."""

from __future__ import absolute_import

import platform
import pytest

from pptx.util import (
    BaseLength, Cm, Collection, Emu, Inches, Mm, Px, to_unicode
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


class TestLengthClasses(TestCase):
    """Test BaseLength, Inches, Cm, Mm, Px, and Emu classes"""
    def test_base_method_values(self):
        """BaseLength() provides correct values for base methods"""
        # setup -----------------------
        expected_px = 96 if platform.system() == 'Windows' else 72
        # exercise --------------------
        x = BaseLength(914400)
        # verify ----------------------
        expected = (1.0, 2.54, 25.4, expected_px, 914400, 914400)
        actual = (x.inches, x.cm, x.mm, x.px, x.emu, x)
        msg = "\nExpected: %s\n     Got: %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_convenience_constructors(self):
        """Length constructors produce a length with correct values"""
        # setup -----------------------
        expected_px = 132 if platform.system() == 'Windows' else 99
        # exercise --------------------
        lengths = {'Inches()': Inches(1.375), 'Cm()': Cm(3.4925),
                   'Mm()': Mm(34.925), 'Px()': Px(expected_px),
                   'Emu()': Emu(1257300)}
        # verify ----------------------
        for constructor, x in lengths.iteritems():
            expected = (1.375, 3.4925, 34.925, expected_px, 1257300, 1257300)
            actual = (x.inches, x.cm, x.mm, x.px, x.emu, x)
            msg = ("for constructor '%s'\nExpected: %s\n     Got: %s"
                   % (constructor, expected, actual))
            self.assertEqual(expected, actual, msg)
