# encoding: utf-8

"""
Test suite for pptx.oxml.text module.
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.text import (
    CT_RegularTextRun, CT_TextBody, CT_TextBodyProperties,
    CT_TextCharacterProperties, CT_TextParagraph, CT_TextParagraphProperties
)

from ..oxml.unitdata.dml import a_gradFill, a_noFill, a_solidFill
from ..oxml.unitdata.text import (
    a_bodyPr, a_defRPr, a_p, a_pPr, a_t, a_txBody, an_endParaRPr, an_extLst,
    an_r, an_rPr
)


class DescribeCT_RegularTextRun(object):

    def it_is_used_by_the_parser_for_an_r_element(self, r):
        assert isinstance(r, CT_RegularTextRun)

    def it_can_get_the_rPr_child_element(self, r_with_rPr, rPr):
        _rPr = r_with_rPr.get_or_add_rPr()
        assert _rPr is rPr

    def it_adds_rPr_element_in_proper_sequence_if_r_doesnt_have_one(
            self, r, r_with_rPr_xml):
        r.get_or_add_rPr()
        assert r.xml == r_with_rPr_xml

    # fixtures ---------------------------------------------

    @pytest.fixture
    def r(self):
        return an_r().with_nsdecls().with_child(a_t()).element

    @pytest.fixture
    def rPr(self):
        return an_rPr().with_nsdecls().element

    @pytest.fixture
    def r_with_rPr(self, r, rPr):
        r.insert(0, rPr)
        return r

    @pytest.fixture
    def r_with_rPr_xml(self):
        rPr_bldr = an_rPr()
        t_bldr = a_t()
        r_bldr = an_r().with_nsdecls()
        r_bldr.with_child(rPr_bldr)
        r_bldr.with_child(t_bldr)
        return r_bldr.xml()


class DescribeCT_TextBody(object):

    def it_is_used_by_the_parser_for_a_txBody_element(self, txBody):
        assert isinstance(txBody, CT_TextBody)

    # fixtures ---------------------------------------------

    @pytest.fixture
    def txBody(self):
        return a_txBody().with_nsdecls().element


class DescribeCT_TextBodyProperties(object):

    def it_is_used_by_the_parser_for_a_bodyPr_element(self, bodyPr):
        assert isinstance(bodyPr, CT_TextBodyProperties)

    def it_can_get_its_xIns_attr_domain_values(self, bodyPr):
        assert bodyPr.lIns is None
        bodyPr.lIns = 123
        assert bodyPr.lIns == 123

        assert bodyPr.tIns is None
        bodyPr.tIns = 456
        assert bodyPr.tIns == 456

        assert bodyPr.rIns is None
        bodyPr.rIns = 789
        assert bodyPr.rIns == 789

        assert bodyPr.bIns is None
        bodyPr.bIns = 876
        assert bodyPr.bIns == 876

    def it_can_set_its_lIns_attribute(
            self, bodyPr, bodyPr_xml, bodyPr_with_lIns_xml,
            bodyPr_with_tIns_xml, bodyPr_with_rIns_xml,
            bodyPr_with_bIns_xml):
        assert bodyPr.xml == bodyPr_xml

        bodyPr.lIns = 987
        assert bodyPr.xml == bodyPr_with_lIns_xml
        bodyPr.lIns = None
        assert bodyPr.xml == bodyPr_xml

        bodyPr.tIns = 654
        assert bodyPr.xml == bodyPr_with_tIns_xml
        bodyPr.tIns = None
        assert bodyPr.xml == bodyPr_xml

        bodyPr.rIns = 321
        assert bodyPr.xml == bodyPr_with_rIns_xml
        bodyPr.rIns = None
        assert bodyPr.xml == bodyPr_xml

        bodyPr.bIns = 234
        assert bodyPr.xml == bodyPr_with_bIns_xml
        bodyPr.bIns = None
        assert bodyPr.xml == bodyPr_xml

    # fixtures ---------------------------------------------

    @pytest.fixture
    def bodyPr(self):
        return a_bodyPr().with_nsdecls().element

    @pytest.fixture
    def bodyPr_xml(self):
        return a_bodyPr().with_nsdecls().xml()

    @pytest.fixture
    def bodyPr_with_lIns_xml(self):
        return a_bodyPr().with_nsdecls().with_lIns(987).xml()

    @pytest.fixture
    def bodyPr_with_tIns_xml(self):
        return a_bodyPr().with_nsdecls().with_tIns(654).xml()

    @pytest.fixture
    def bodyPr_with_rIns_xml(self):
        return a_bodyPr().with_nsdecls().with_rIns(321).xml()

    @pytest.fixture
    def bodyPr_with_bIns_xml(self):
        return a_bodyPr().with_nsdecls().with_bIns(234).xml()


class DescribeCT_TextCharacterProperties(object):

    def it_is_used_by_the_parser_for_a_defRPr_element(self, defRPr):
        assert isinstance(defRPr, CT_TextCharacterProperties)

    def it_is_used_by_the_parser_for_an_endParaRPr_element(self, endParaRPr):
        assert isinstance(endParaRPr, CT_TextCharacterProperties)

    def it_is_used_by_the_parser_for_an_rPr_element(self, rPr):
        assert isinstance(rPr, CT_TextCharacterProperties)

    def it_knows_the_b_value(self, rPr_with_true_b, rPr_with_false_b, rPr):
        assert rPr_with_true_b.b is True
        assert rPr_with_false_b.b is False
        assert rPr.b is None

    def it_can_set_the_b_value(
            self, rPr, rPr_with_true_b_xml, rPr_with_false_b_xml, rPr_xml):
        rPr.b = True
        assert rPr.xml == rPr_with_true_b_xml
        rPr.b = False
        assert rPr.xml == rPr_with_false_b_xml
        rPr.b = None
        assert rPr.xml == rPr_xml

    def it_knows_the_i_value(self, rPr_with_true_i, rPr_with_false_i, rPr):
        assert rPr_with_true_i.i is True
        assert rPr_with_false_i.i is False
        assert rPr.i is None

    def it_can_set_the_i_value(
            self, rPr, rPr_with_true_i_xml, rPr_with_false_i_xml, rPr_xml):
        rPr.i = True
        assert rPr.xml == rPr_with_true_i_xml
        rPr.i = False
        assert rPr.xml == rPr_with_false_i_xml
        rPr.i = None
        assert rPr.xml == rPr_xml

    def it_can_get_the_solidFill_child_element_or_none_if_there_isnt_one(
            self, rPr, rPr_with_solidFill, solidFill):
        assert rPr.solidFill is None
        assert rPr_with_solidFill.solidFill is solidFill

    def it_gets_the_solidFill_child_element_if_there_is_one(
            self, rPr_with_solidFill, solidFill):
        _solidFill = rPr_with_solidFill.get_or_change_to_solidFill()
        assert _solidFill is solidFill

    def it_adds_a_solidFill_child_element_if_there_isnt_one(
            self, rPr, rPr_with_solidFill_xml):
        solidFill = rPr.get_or_change_to_solidFill()
        assert rPr.xml == rPr_with_solidFill_xml
        assert rPr.find(qn('a:solidFill')) == solidFill

    def it_changes_the_fill_type_to_solidFill_if_another_one_is_there(
            self, rPr_with_gradFill, rPr_with_noFill, rPr_with_solidFill_xml):
        rPr_with_gradFill.get_or_change_to_solidFill()
        assert rPr_with_gradFill.xml == rPr_with_solidFill_xml
        rPr_with_noFill.get_or_change_to_solidFill()
        assert rPr_with_noFill.xml == rPr_with_solidFill_xml

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

    @pytest.fixture
    def rPr_xml(self):
        return an_rPr().with_nsdecls().xml()

    @pytest.fixture
    def rPr_with_false_b(self):
        return an_rPr().with_nsdecls().with_b('false').element

    @pytest.fixture
    def rPr_with_false_b_xml(self):
        return an_rPr().with_nsdecls().with_b(0).xml()

    @pytest.fixture
    def rPr_with_false_i(self):
        return an_rPr().with_nsdecls().with_i('false').element

    @pytest.fixture
    def rPr_with_false_i_xml(self):
        return an_rPr().with_nsdecls().with_i(0).xml()

    @pytest.fixture
    def rPr_with_gradFill(self):
        gradFill_bldr = a_gradFill()
        return an_rPr().with_nsdecls().with_child(gradFill_bldr).element

    @pytest.fixture
    def rPr_with_noFill(self):
        noFill_bldr = a_noFill()
        return an_rPr().with_nsdecls().with_child(noFill_bldr).element

    @pytest.fixture
    def rPr_with_solidFill(self, solidFill):
        rPr = an_rPr().with_nsdecls().element
        rPr.append(solidFill)
        return rPr

    @pytest.fixture
    def rPr_with_solidFill_xml(self):
        solidFill_bldr = a_solidFill()
        return an_rPr().with_nsdecls().with_child(solidFill_bldr).xml()

    @pytest.fixture
    def rPr_with_true_b(self):
        return an_rPr().with_nsdecls().with_b(1).element

    @pytest.fixture
    def rPr_with_true_b_xml(self):
        return an_rPr().with_nsdecls().with_b(1).xml()

    @pytest.fixture
    def rPr_with_true_i(self):
        return an_rPr().with_nsdecls().with_i(1).element

    @pytest.fixture
    def rPr_with_true_i_xml(self):
        return an_rPr().with_nsdecls().with_i(1).xml()

    @pytest.fixture
    def solidFill(self):
        return a_solidFill().with_nsdecls().element


class DescribeCT_TextParagraph(object):

    def it_is_used_by_the_parser_for_a_p_element(self, p):
        assert isinstance(p, CT_TextParagraph)

    def it_can_get_the_pPr_child_element(self, p_with_pPr, pPr):
        _pPr = p_with_pPr.get_or_add_pPr()
        assert _pPr is pPr

    def it_adds_a_pPr_if_p_doesnt_have_one(self, p, p_with_pPr_xml):
        p.get_or_add_pPr()
        assert p.xml == p_with_pPr_xml

    def it_can_add_a_new_r_element(self, p, p_with_r_xml):
        p.add_r()
        assert p.xml == p_with_r_xml

    def it_adds_r_element_in_correct_sequence(
            self, p_with_endParaRPr, p_with_r_with_endParaRPr_xml):
        p = p_with_endParaRPr
        p.add_r()
        assert p.xml == p_with_r_with_endParaRPr_xml

    def it_can_remove_all_its_r_child_elements(
            self, p_with_r_children, p_xml):
        p = p_with_r_children.remove_child_r_elms()
        assert p.xml == p_xml

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
        assert pPr_with_algn.algn == PP_ALIGN.THAI_DISTRIBUTE

    def it_maps_missing_algn_attribute_to_None(self, pPr):
        assert pPr.algn is None

    def it_can_set_the_algn_value(self, pPr, pPr_with_algn_xml, pPr_xml):
        pPr.algn = PP_ALIGN.THAI_DISTRIBUTE
        assert pPr.xml == pPr_with_algn_xml
        pPr.algn = None
        assert pPr.xml == pPr_xml

    def it_can_get_the_defRPr_child_element(self, pPr_with_defRPr, defRPr):
        _defRPr = pPr_with_defRPr.get_or_add_defRPr()
        assert _defRPr is defRPr

    def it_adds_a_defRPr_if_pPr_doesnt_have_one(
            self, pPr, pPr_with_defRPr_xml):
        pPr.get_or_add_defRPr()
        assert pPr.xml == pPr_with_defRPr_xml

    def it_adds_defRPr_element_in_correct_sequence(
            self, pPr_with_extLst, pPr_with_defRPr_with_extLst_xml):
        pPr = pPr_with_extLst
        pPr.get_or_add_defRPr()
        assert pPr.xml == pPr_with_defRPr_with_extLst_xml

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
        return a_pPr().with_nsdecls().with_algn('thaiDist')

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
