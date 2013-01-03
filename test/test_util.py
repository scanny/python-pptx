# -*- coding: utf-8 -*-
#
# test_util.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Test suite for pptx.util module."""

import pptx.util

from testing import TestCase

class TestPartname(TestCase):
    """Test pptx.util.Partname"""
    def setUp(self):
        self.partname_str = '/ppt/slides/slide23.xml'
        self.baseURI = '/ppt/slides'
        self.filename = 'slide23.xml'
        self.ext = '.xml'
        self.basename = 'slide'
        self.idx = 23
        self.partname = pptx.util.Partname(self.partname_str)
    
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
        partname = pptx.util.Partname('/ppt/presentation.xml')
        # exercise --------------------
        retval = partname.idx
        # verify ----------------------
        expected = None
        actual = retval
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

