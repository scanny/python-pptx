# encoding: utf-8

"""Test suite for pptx.spec module."""

from __future__ import absolute_import

from hamcrest import assert_that, equal_to, is_

from pptx.constants import PP, TEXT_ALIGN_TYPE as TAT
from pptx.spec import ParagraphAlignment

from .unitutil import TestCase


class TestParagraphAlignment(TestCase):
    """Test ParagraphAlignment"""
    cases = (
        (PP.ALIGN_CENTER,     TAT.CENTER),
        (PP.ALIGN_DISTRIBUTE, TAT.DISTRIBUTE),
        (PP.ALIGN_JUSTIFY,    TAT.JUSTIFY),
        (None,                None)
    )

    def test_from_text_align_type_return_value(self):
        """ParagraphAlignment.from_text_align_type() returns correct value"""
        # verify -----------------------
        for alignment, text_align_type in self.cases:
            assert_that(
                ParagraphAlignment.from_text_align_type(text_align_type),
                is_(equal_to(alignment))
            )

    def test_from_text_align_type_raises_on_bad_key(self):
        """ParagraphAlignment.from_text_align_type raises on bad key"""
        with self.assertRaises(KeyError):
            ParagraphAlignment.from_text_align_type('foobar')

    def test_to_text_align_type_return_value(self):
        """ParagraphAlignment.to_text_align_type returns correct value"""
        # verify -----------------------
        for alignment, text_align_type in self.cases:
            assert_that(ParagraphAlignment.to_text_align_type(alignment),
                        is_(equal_to(text_align_type)))

    def test_to_text_align_type_raises_on_bad_alignment(self):
        """ParagraphAlignment.to_text_align_type raises on bad value in"""
        with self.assertRaises(KeyError):
            ParagraphAlignment.to_text_align_type('foobar')
