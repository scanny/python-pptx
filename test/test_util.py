# -*- coding: utf-8 -*-
#
# test_util.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Test suite for pptx.util module."""

import platform

from pptx.util import BaseLength, Cm, Emu, Inches, Mm, Partname, Px

from testing import TestCase

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
        lengths =\
            { 'Inches()' : Inches(1.375)
            , 'Cm()'     : Cm(3.4925)
            , 'Mm()'     : Mm(34.925)
            , 'Px()'     : Px(expected_px)
            , 'Emu()'    : Emu(1257300)
            }
        # verify ----------------------
        for constructor, x in lengths.iteritems():
            expected = (1.375, 3.4925, 34.925, expected_px, 1257300, 1257300)
            actual = (x.inches, x.cm, x.mm, x.px, x.emu, x)
            msg = ("for constructor '%s'\nExpected: %s\n     Got: %s"
                   % (constructor, expected, actual))
            self.assertEqual(expected, actual, msg)
    

class TestPartname(TestCase):
    """Test pptx.util.Partname"""
    def setUp(self):
        self.partname_str = '/ppt/slides/slide23.xml'
        self.baseURI = '/ppt/slides'
        self.filename = 'slide23.xml'
        self.ext = '.xml'
        self.basename = 'slide'
        self.idx = 23
        self.partname = Partname(self.partname_str)
    
    def test_partname_correct(self):
        """Partname.partname contains correct partname string"""
        # exercise --------------------
        retval = self.partname.partname
        # verify ----------------------
        expected = self.partname_str
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_baseuri_correct(self):
        """Partname.baseURI value correct"""
        # exercise --------------------
        retval = self.partname.baseURI
        # verify ----------------------
        expected = self.baseURI
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_filename_correct(self):
        """Partname.filename value correct"""
        # exercise --------------------
        retval = self.partname.filename
        # verify ----------------------
        expected = self.filename
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_ext_correct(self):
        """Partname.ext value correct"""
        # exercise --------------------
        retval = self.partname.ext
        # verify ----------------------
        expected = self.ext
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_basename_correct(self):
        """Partname.basename value correct"""
        # exercise --------------------
        retval = self.partname.basename
        # verify ----------------------
        expected = self.basename
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_idx_correct_for_tuple_part(self):
        """Partname.idx value correct for tuple partname"""
        # exercise --------------------
        retval = self.partname.idx
        # verify ----------------------
        expected = self.idx
        actual = retval
        msg = "expected %d, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_idx_correct_for_singleton_part(self):
        """Partname.idx value correct for singleton partname"""
        # setup -----------------------
        partname = Partname('/ppt/presentation.xml')
        # exercise --------------------
        retval = partname.idx
        # verify ----------------------
        expected = None
        actual = retval
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

