# encoding: utf-8

"""Test suite for pptx.oxml module."""

from __future__ import absolute_import

from hamcrest import assert_that, equal_to, is_

from pptx.constants import TEXT_ALIGN_TYPE as TAT

from ..oxml.unitdata.text import test_text_elements, test_text_xml
from ..unitutil import TestCase


class TestCT_TextParagraph(TestCase):
    """Test CT_TextParagraph"""
    def test_get_algn_returns_correct_value(self):
        """CT_TextParagraph.get_algn() returns correct value"""
        # setup ------------------------
        cases = (
            (test_text_elements.paragraph, None),
            (test_text_elements.centered_paragraph, TAT.CENTER)
        )
        # verify -----------------------
        for p, expected_algn in cases:
            assert_that(p.get_algn(), is_(equal_to(expected_algn)))

    def test_set_algn_sets_algn_value(self):
        """CT_TextParagraph.set_algn() sets algn value"""
        # setup ------------------------
        cases = (
            # something => something else
            (test_text_elements.centered_paragraph, TAT.JUSTIFY),
            # something => None
            (test_text_elements.centered_paragraph, None),
            # None => something
            (test_text_elements.paragraph, TAT.CENTER),
            # None => None
            (test_text_elements.paragraph, None)
        )
        # verify -----------------------
        for p, algn in cases:
            p.set_algn(algn)
            assert_that(p.get_algn(), is_(equal_to(algn)))

    def test_set_algn_produces_correct_xml(self):
        """Assigning value to CT_TextParagraph.algn produces correct XML"""
        # setup ------------------------
        cases = (
            # None => something
            (test_text_elements.paragraph, TAT.CENTER,
             test_text_xml.centered_paragraph),
            # something => None
            (test_text_elements.centered_paragraph, None,
             test_text_xml.paragraph)
        )
        # verify -----------------------
        for p, text_align_type, expected_xml in cases:
            p.set_algn(text_align_type)
            self.assertEqualLineByLine(expected_xml, p)
