# encoding: utf-8

"""Test suite for pptx.oxml module."""

from __future__ import absolute_import

import pytest

from pptx.constants import TEXT_ALIGN_TYPE as TAT
from pptx.oxml.text import CT_TextParagraph, CT_TextParagraphProperties

from ..oxml.unitdata.text import (
    a_p, a_pPr, a_t, an_endParaRPr, an_r, test_text_elements, test_text_xml
)
from ..unitutil import actual_xml


class DescribeCT_TextParagraph(object):

    def it_is_used_by_the_parser_for_a_p_element(self, p):
        assert isinstance(p, CT_TextParagraph)

    def it_can_get_the_pPr_child_element(self, p_with_pPr, pPr):
        _pPr = p_with_pPr.get_or_add_pPr()
        assert _pPr is pPr

    def it_adds_a_pPr_if_p_doesnt_have_one(self, p, p_with_pPr_xml):
        p.get_or_add_pPr()
        assert actual_xml(p) == p_with_pPr_xml

    def it_can_add_a_new_r_element(self, p, p_with_r_xml):
        p.add_r()
        assert actual_xml(p) == p_with_r_xml

    def it_adds_r_element_in_correct_sequence(
            self, p_with_endParaRPr, p_with_r_with_endParaRPr_xml):
        p = p_with_endParaRPr
        p.add_r()
        assert actual_xml(p) == p_with_r_with_endParaRPr_xml

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
            assert p.get_or_add_pPr().algn == algn

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
            assert actual_xml(p) == expected_xml

    # fixtures ---------------------------------------------

    @pytest.fixture
    def p(self):
        p_bldr = a_p().with_nsdecls()
        return p_bldr.element

    @pytest.fixture
    def pPr(self):
        return a_pPr().with_nsdecls().element

    @pytest.fixture
    def p_with_r_xml(self):
        r_bldr = an_r().with_child(a_t())
        return a_p().with_nsdecls().with_child(r_bldr).xml()

    @pytest.fixture
    def p_with_endParaRPr(self):
        endParaRPr_bldr = an_endParaRPr()
        p_bldr = a_p().with_nsdecls().with_child(endParaRPr_bldr)
        return p_bldr.element

    @pytest.fixture
    def p_with_pPr(self, p, pPr):
        p.append(pPr)
        return p

    @pytest.fixture
    def p_with_pPr_xml(self):
        pPr_bldr = a_pPr()
        p_with_pPr_bldr = a_p().with_nsdecls().with_child(pPr_bldr)
        return p_with_pPr_bldr.xml()

    @pytest.fixture
    def p_with_r_with_endParaRPr_xml(self):
        r_bldr = an_r().with_child(a_t())
        endParaRPr_bldr = an_endParaRPr()
        p_bldr = a_p().with_nsdecls()
        p_bldr = p_bldr.with_child(r_bldr)
        p_bldr = p_bldr.with_child(endParaRPr_bldr)
        return p_bldr.xml()


class DescribeCT_TextParagraphProperties(object):

    def it_is_used_by_the_parser_for_a_pPr_element(self, pPr):
        assert isinstance(pPr, CT_TextParagraphProperties)

    def it_knows_the_algn_value(self, pPr_with_algn):
        assert pPr_with_algn.algn == 'foobar'

    def it_maps_missing_algn_attribute_to_None(self, pPr):
        assert pPr.algn is None

    # fixtures ---------------------------------------------

    @pytest.fixture
    def pPr(self):
        pPr_bldr = a_pPr().with_nsdecls()
        return pPr_bldr.element

    @pytest.fixture
    def pPr_with_algn(self):
        return a_pPr().with_nsdecls().with_algn('foobar').element
