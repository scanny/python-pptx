# -*- coding: utf-8 -*-
#
# test_spec.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under the MIT License:
# http://www.opensource.org/licenses/mit-license.php

"""Test suite for pptx.spec module."""

from hamcrest import assert_that, equal_to, is_

from pptx.constants import PP, TEXT_ALIGN_TYPE as TAT
from pptx.spec import ParagraphAlignment

from testing import TestCase


class TestParagraphAlignment(TestCase):
    """Test ParagraphAlignment"""
    def test_to_text_align_type_return_value(self):
        """ParagraphAlignment.to_text_align_type returns correct value"""
        # setup ------------------------
        cases = (
            (PP.ALIGN_CENTER,     TAT.CENTER),
            (PP.ALIGN_DISTRIBUTE, TAT.DISTRIBUTE),
            (PP.ALIGN_JUSTIFY,    TAT.JUSTIFY)
        )
        # verify -----------------------
        for alignment, text_align_type in cases:
            assert_that(ParagraphAlignment.to_text_align_type(alignment),
                        is_(equal_to(text_align_type)))

    def test_to_text_align_type_raises_on_bad_alignment(self):
        """ParagraphAlignment.to_text_align_type raises on bad value in"""
        with self.assertRaises(KeyError):
            ParagraphAlignment.to_text_align_type('foobar')
