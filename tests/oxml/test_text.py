# encoding: utf-8

"""Test suite for pptx.oxml module."""

from __future__ import absolute_import

import pytest

from pptx.oxml.text import (
    CT_TextCharacterProperties, CT_TextParagraph, CT_TextParagraphProperties
)

from ..oxml.unitdata.text import (
    a_defRPr, a_p, a_pPr, a_t, an_endParaRPr, an_extLst, an_r, an_rPr
)
from ..unitutil import actual_xml


class DescribeCT_TextCharacterProperties(object):

    def it_is_used_by_the_parser_for_a_defRPr_element(self, defRPr):
        assert isinstance(defRPr, CT_TextCharacterProperties)

    def it_is_used_by_the_parser_for_an_endParaRPr_element(self, endParaRPr):
        assert isinstance(endParaRPr, CT_TextCharacterProperties)

    def it_is_used_by_the_parser_for_an_rPr_element(self, rPr):
        assert isinstance(rPr, CT_TextCharacterProperties)

    # fixtures ---------------------------------------------

    @pytest.fixture
    def defRPr(self):
        return a_defRPr().with_nsdecls().element

    @pytest.fixture
    def endParaRPr(self):
        return an_endParaRPr().with_nsdecls().element

    @pytest.fixture
    def rPr(self):
        return an_rPr().with_nsdecls().element


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

    def it_can_remove_all_its_r_child_elements(
            self, p_with_r_children, p_xml):
        p = p_with_r_children.remove_child_r_elms()
        assert actual_xml(p) == p_xml

    # fixtures ---------------------------------------------

    @pytest.fixture
    def p(self, p_bldr):
        return p_bldr.element

    @pytest.fixture
    def p_bldr(self):
        return a_p().with_nsdecls()

    @pytest.fixture
    def p_xml(self, p_bldr):
        return p_bldr.xml()

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
    def p_with_r_children(self):
        r_bldr = an_r().with_child(a_t())
        p_bldr = a_p().with_nsdecls()
        p_bldr = p_bldr.with_child(r_bldr)
        p_bldr = p_bldr.with_child(r_bldr)
        return p_bldr.element

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

    def it_can_set_the_algn_value(self, pPr, pPr_with_algn_xml, pPr_xml):
        pPr.algn = 'foobar'
        assert actual_xml(pPr) == pPr_with_algn_xml
        pPr.algn = None
        assert actual_xml(pPr) == pPr_xml

    def it_can_get_the_defRPr_child_element(self, pPr_with_defRPr, defRPr):
        _defRPr = pPr_with_defRPr.get_or_add_defRPr()
        assert _defRPr is defRPr

    def it_adds_a_defRPr_if_pPr_doesnt_have_one(
            self, pPr, pPr_with_defRPr_xml):
        pPr.get_or_add_defRPr()
        assert actual_xml(pPr) == pPr_with_defRPr_xml

    def it_adds_defRPr_element_in_correct_sequence(
            self, pPr_with_extLst, pPr_with_defRPr_with_extLst_xml):
        pPr = pPr_with_extLst
        pPr.get_or_add_defRPr()
        assert actual_xml(pPr) == pPr_with_defRPr_with_extLst_xml

    # fixtures ---------------------------------------------

    @pytest.fixture
    def defRPr(self):
        return a_defRPr().with_nsdecls().element

    @pytest.fixture
    def pPr(self, pPr_bldr):
        return pPr_bldr.element

    @pytest.fixture
    def pPr_bldr(self):
        return a_pPr().with_nsdecls()

    @pytest.fixture
    def pPr_xml(self, pPr_bldr):
        return pPr_bldr.xml()

    @pytest.fixture
    def pPr_with_algn(self, pPr_with_algn_bldr):
        return pPr_with_algn_bldr.element

    @pytest.fixture
    def pPr_with_algn_bldr(self):
        return a_pPr().with_nsdecls().with_algn('foobar')

    @pytest.fixture
    def pPr_with_algn_xml(self, pPr_with_algn_bldr):
        return pPr_with_algn_bldr.xml()

    @pytest.fixture
    def pPr_with_defRPr(self, pPr_bldr, defRPr):
        pPr = pPr_bldr.element
        pPr.append(defRPr)
        return pPr

    @pytest.fixture
    def pPr_with_defRPr_with_extLst_xml(self):
        defRPr_bldr = a_defRPr()
        extLst_bldr = an_extLst()
        pPr_bldr = a_pPr().with_nsdecls()
        pPr_bldr = pPr_bldr.with_child(defRPr_bldr)
        pPr_bldr = pPr_bldr.with_child(extLst_bldr)
        return pPr_bldr.xml()

    @pytest.fixture
    def pPr_with_defRPr_xml(self):
        defRPr_bldr = a_defRPr()
        pPr_bldr = a_pPr().with_nsdecls().with_child(defRPr_bldr)
        return pPr_bldr.xml()

    @pytest.fixture
    def pPr_with_extLst(self):
        extLst_bldr = an_extLst()
        pPr_bldr = a_pPr().with_nsdecls().with_child(extLst_bldr)
        return pPr_bldr.element
