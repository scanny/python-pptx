"""Unit-test suite for `pptx.oxml.timing` module (Foundation F8 MVP)."""

from __future__ import annotations

import pytest

from pptx.oxml.timing import (
    CT_SlideTiming,
    CT_TimeNodeList,
    CT_TLCommonTimeNodeData,
    CT_TLTimeCondition,
    CT_TLTimeConditionList,
    CT_TLTimeNodeParallel,
    CT_TLTimeNodeSequence,
    ST_TLTime,
)

from ..unitutil.cxml import element


class DescribeST_TLTime(object):
    """Unit-test suite for the ST_TLTime simple type."""

    def it_converts_indefinite_to_indefinite(self):
        assert ST_TLTime.convert_from_xml("indefinite") == "indefinite"

    def it_converts_integer_strings_to_int(self):
        assert ST_TLTime.convert_from_xml("500") == 500

    def it_converts_indefinite_back_to_string(self):
        assert ST_TLTime.convert_to_xml("indefinite") == "indefinite"

    def it_converts_int_back_to_string(self):
        assert ST_TLTime.convert_to_xml(1200) == "1200"

    def it_accepts_indefinite_on_validate(self):
        # -- raises nothing --
        ST_TLTime.validate("indefinite")

    def it_accepts_non_negative_ints_on_validate(self):
        ST_TLTime.validate(0)
        ST_TLTime.validate(500)

    def it_rejects_negative_ints_on_validate(self):
        with pytest.raises(ValueError, match="non-negative"):
            ST_TLTime.validate(-1)

    def it_rejects_non_int_non_indefinite_on_validate(self):
        with pytest.raises(ValueError, match="non-negative"):
            ST_TLTime.validate("one second")


class DescribeCT_SlideTiming(object):
    """Unit-test suite for `pptx.oxml.timing.CT_SlideTiming`."""

    def it_is_the_class_used_for_the_p_timing_element(self):
        timing = element("p:timing")
        assert isinstance(timing, CT_SlideTiming)

    def it_exposes_an_optional_tnLst_child(self):
        timing = element("p:timing/p:tnLst")
        assert timing.tnLst is not None

    def and_tnLst_is_None_when_absent(self):
        timing = element("p:timing")
        assert timing.tnLst is None

    def it_can_add_a_tnLst_child(self):
        timing = element("p:timing")
        tnLst = timing.get_or_add_tnLst()
        assert tnLst is timing.tnLst


class DescribeCT_TimeNodeList(object):
    """Unit-test suite for `pptx.oxml.timing.CT_TimeNodeList`."""

    def it_is_used_for_p_tnLst(self):
        tnLst = element("p:tnLst")
        assert isinstance(tnLst, CT_TimeNodeList)

    def it_is_used_for_p_childTnLst(self):
        childTnLst = element("p:childTnLst")
        assert isinstance(childTnLst, CT_TimeNodeList)

    def it_allocates_cTn_id_1_on_empty_tnLst(self):
        # -- Contrived: tnLst not attached to a p:sld, but _next_cTn_id must
        # -- still return 1 as the initial id rather than raising.
        tnLst = element("p:tnLst")
        assert tnLst._next_cTn_id == 1


class DescribeCT_TLCommonTimeNodeData(object):
    """Unit-test suite for `pptx.oxml.timing.CT_TLCommonTimeNodeData`."""

    def it_is_used_for_the_p_cTn_element(self):
        cTn = element("p:cTn")
        assert isinstance(cTn, CT_TLCommonTimeNodeData)

    def it_reads_its_id_attribute(self):
        cTn = element('p:cTn{id=3}')
        assert cTn.id == 3

    def and_id_is_None_when_absent(self):
        cTn = element("p:cTn")
        assert cTn.id is None

    def it_reads_its_nodeType_attribute(self):
        cTn = element('p:cTn{nodeType=tmRoot}')
        assert cTn.nodeType == "tmRoot"

    def it_reads_its_dur_attribute_as_int(self):
        cTn = element('p:cTn{dur=500}')
        assert cTn.dur == 500

    def and_reads_indefinite_as_string(self):
        cTn = element('p:cTn{dur=indefinite}')
        assert cTn.dur == "indefinite"

    def and_dur_is_None_when_absent(self):
        cTn = element("p:cTn")
        assert cTn.dur is None

    def it_exposes_an_optional_stCondLst_child(self):
        cTn = element("p:cTn/p:stCondLst")
        assert cTn.stCondLst is not None

    def and_stCondLst_is_None_when_absent(self):
        cTn = element("p:cTn")
        assert cTn.stCondLst is None

    def it_can_get_or_add_a_stCondLst_child(self):
        cTn = element("p:cTn")
        stCondLst = cTn.get_or_add_stCondLst()
        assert isinstance(stCondLst, CT_TLTimeConditionList)
        assert cTn.stCondLst is stCondLst


class DescribeCT_TLTimeConditionList(object):
    """Unit-test suite for `pptx.oxml.timing.CT_TLTimeConditionList`."""

    def it_is_used_for_the_p_stCondLst_element(self):
        stCondLst = element("p:stCondLst")
        assert isinstance(stCondLst, CT_TLTimeConditionList)

    def it_is_used_for_the_p_endCondLst_element(self):
        endCondLst = element("p:endCondLst")
        assert isinstance(endCondLst, CT_TLTimeConditionList)

    def it_can_add_a_cond_child(self):
        stCondLst = element("p:stCondLst")
        cond = stCondLst.add_cond()
        assert isinstance(cond, CT_TLTimeCondition)


class DescribeCT_TLTimeCondition(object):
    """Unit-test suite for `pptx.oxml.timing.CT_TLTimeCondition`."""

    def it_is_used_for_the_p_cond_element(self):
        cond = element("p:cond")
        assert isinstance(cond, CT_TLTimeCondition)

    def it_reads_its_evt_attribute(self):
        cond = element("p:cond{evt=onClick}")
        assert cond.evt == "onClick"

    def and_evt_is_None_when_absent(self):
        cond = element("p:cond")
        assert cond.evt is None

    def it_reads_its_delay_as_int(self):
        cond = element("p:cond{delay=1500}")
        assert cond.delay == 1500

    def and_reads_indefinite_as_string(self):
        cond = element("p:cond{delay=indefinite}")
        assert cond.delay == "indefinite"

    def and_delay_is_None_when_absent(self):
        cond = element("p:cond")
        assert cond.delay is None


class DescribeCT_TLTimeNodeParallel(object):
    """Unit-test suite for `pptx.oxml.timing.CT_TLTimeNodeParallel`."""

    def it_is_used_for_the_p_par_element(self):
        par = element("p:par/p:cTn")
        assert isinstance(par, CT_TLTimeNodeParallel)

    def it_exposes_its_required_cTn_child(self):
        par = element("p:par/p:cTn{id=1}")
        assert par.cTn.id == 1


class DescribeCT_TLTimeNodeSequence(object):
    """Unit-test suite for `pptx.oxml.timing.CT_TLTimeNodeSequence`."""

    def it_is_used_for_the_p_seq_element(self):
        seq = element("p:seq/p:cTn")
        assert isinstance(seq, CT_TLTimeNodeSequence)

    def it_exposes_its_required_cTn_child(self):
        seq = element("p:seq/p:cTn{id=2,nodeType=mainSeq}")
        assert seq.cTn.nodeType == "mainSeq"
