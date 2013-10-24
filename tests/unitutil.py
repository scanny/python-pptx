# -*- coding: utf-8 -*-
#
# testing.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Testing utilities for python-pptx."""

import unittest2

from pptx.oxml import oxml_tostring


class TestCase(unittest2.TestCase):
    """Additional assert methods for python-pptx unit testing."""
    def assertEqualLineByLine(self, expected_xml, element):
        """
        Apply assertEqual() to each line of *expected_xml* and corresponding
        line of XML derived from *element*.
        """
        actual_xml = oxml_tostring(element, pretty_print=True)
        actual_xml_lines = actual_xml.split('\n')
        expected_xml_lines = expected_xml.split('\n')
        for idx, line in enumerate(actual_xml_lines):
            msg = ("\n\nexpected:\n\n%s'\nbut got\n\n%s'" %
                   (expected_xml, actual_xml))
            self.assertEqual(line, expected_xml_lines[idx], msg)

    def assertIsInstance(self, obj, cls):
        """Raise AssertionError if *obj* is not instance of *cls*."""
        tmpl = "expected instance of '%s', got type '%s'"
        if not isinstance(obj, cls):
            raise AssertionError(tmpl % (cls.__name__, type(obj).__name__))

    def assertIsProperty(self, inst, propname, value, read_only=True):
        """
        Raise AssertionError if *propname* is not a property of *obj* having
        the specified characteristics. Will raise AssertionError if *propname*
        does not exist in *obj*, if its value is not equal to *value*, or if
        *read_only* is True and assignment does not raise AttributeError.
        """
        if not hasattr(inst, propname):
            tmpl = "expected %s to have attribute '%s'"
            raise AssertionError(tmpl % (inst, propname))
        expected = value
        actual = getattr(inst, propname)
        if actual != expected:
            tmpl = "expected '%s', got '%s'"
            raise AssertionError(tmpl % (expected, actual))
        if read_only:
            try:
                with self.assertRaises(AttributeError):
                    setattr(inst, propname, None)
            except AssertionError:
                tmpl = "property '%s' on class '%s' is not read-only"
                clsname = inst.__class__.__name__
                raise AssertionError(tmpl % (propname, clsname))

    def assertIsReadOnly(self, inst, propname):
        """
        Raise AssertionError if *propname* does not raise AttributeError when
        assignment is attempted.
        """
        try:
            with self.assertRaises(AttributeError):
                setattr(inst, propname, None)
        except AssertionError:
            tmpl = "%s.%s is not read-only"
            clsname = inst.__class__.__name__
            raise AssertionError(tmpl % (clsname, propname))

    def assertIsSizedProperty(self, inst, propname, length, read_only=True):
        """
        Raise AssertionError if *propname* is not a property of *obj* having
        the specified characteristics. Will raise AssertionError if *propname*
        does not exist in *obj*, if len(inst.propname) is not equal to
        *length*, or if *read_only* is True and assignment does not raise
        AttributeError.
        """
        if not hasattr(inst, propname):
            tmpl = "expected %s to have attribute '%s'"
            raise AssertionError(tmpl % (inst, propname))
        expected = length
        actual = len(getattr(inst, propname))
        if actual != expected:
            tmpl = "expected length %d, got %d"
            raise AssertionError(tmpl % (expected, actual))
        if read_only:
            try:
                with self.assertRaises(AttributeError):
                    setattr(inst, propname, None)
            except AssertionError:
                tmpl = "property '%s' on class '%s' is not read-only"
                clsname = inst.__class__.__name__
                raise AssertionError(tmpl % (propname, clsname))

    def assertLength(self, sized, length):
        """Raise AssertionError if len(*sized*) != *length*"""
        expected = length
        actual = len(sized)
        msg = "expected length %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
