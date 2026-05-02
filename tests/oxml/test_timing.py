"""Unit-test suite for `pptx.oxml.timing` module (Foundation F8 MVP)."""

from __future__ import annotations

import pytest

from pptx.oxml.timing import (
    CT_SlideTiming,
    CT_TimeNodeList,
    CT_TLCommonTimeNodeData,
    CT_TLShapeTargetElement,
    CT_TLTimeCondition,
    CT_TLTimeConditionList,
    CT_TLTimeNodeParallel,
    CT_TLTimeNodeSequence,
    ST_TLTime,
    first_spTgt_spid,
    iter_main_sequence_effects,
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

    # -- stCondLst / endCondLst descriptors (#861) -------------------


    def it_exposes_an_optional_stCondLst_child(self):
        cTn = element("p:cTn/p:stCondLst")
        assert cTn.stCondLst is not None

    def and_stCondLst_is_None_when_absent(self):
        cTn = element("p:cTn")
        assert cTn.stCondLst is None

    def it_exposes_an_optional_endCondLst_child(self):
        cTn = element("p:cTn/p:endCondLst")
        assert cTn.endCondLst is not None

    def it_can_get_or_add_a_stCondLst_child(self):
        cTn = element("p:cTn")
        stCondLst = cTn.get_or_add_stCondLst()
        assert isinstance(stCondLst, CT_TLTimeConditionList)
        assert cTn.stCondLst is stCondLst

    # -- delay read ---------------------------------------------------

    def it_reads_delay_from_the_first_stCondLst_cond(self):
        cTn = element("p:cTn/p:stCondLst/p:cond{delay=500}")
        assert cTn.delay == 500

    def and_reads_indefinite_delay_as_string(self):
        cTn = element("p:cTn/p:stCondLst/p:cond{delay=indefinite}")
        assert cTn.delay == "indefinite"

    def and_delay_is_None_when_stCondLst_is_absent(self):
        cTn = element("p:cTn")
        assert cTn.delay is None

    def and_delay_is_None_when_stCondLst_has_no_cond(self):
        cTn = element("p:cTn/p:stCondLst")
        assert cTn.delay is None

    def and_delay_is_None_when_cond_has_no_delay_attr(self):
        cTn = element("p:cTn/p:stCondLst/p:cond{evt=onClick}")
        assert cTn.delay is None

    # -- delay write --------------------------------------------------

    def it_writes_delay_by_adding_stCondLst_and_cond(self):
        cTn = element("p:cTn")
        cTn.delay = 1000
        assert cTn.delay == 1000
        assert cTn.stCondLst is not None

    def and_writes_indefinite_delay(self):
        cTn = element("p:cTn")
        cTn.delay = "indefinite"
        assert cTn.delay == "indefinite"

    def and_overwrites_an_existing_delay_without_adding_a_new_cond(self):
        cTn = element("p:cTn/p:stCondLst/p:cond{delay=500,evt=onClick}")
        cTn.delay = 2000
        assert cTn.delay == 2000
        # -- @evt on the existing cond is preserved (we mutated, not appended) --
        from pptx.oxml.ns import qn

        conds = cTn.stCondLst.findall(qn("p:cond"))
        assert len(conds) == 1
        assert conds[0].evt == "onClick"

    def and_setting_None_clears_the_delay_attribute_in_place(self):
        cTn = element("p:cTn/p:stCondLst/p:cond{delay=500,evt=onClick}")
        cTn.delay = None
        assert cTn.delay is None
        # -- cond itself is retained so caller-set @evt survives --
        from pptx.oxml.ns import qn

        conds = cTn.stCondLst.findall(qn("p:cond"))
        assert len(conds) == 1
        assert conds[0].evt == "onClick"

    def and_setting_None_is_a_no_op_when_no_stCondLst_is_present(self):
        cTn = element("p:cTn")
        cTn.delay = None
        assert cTn.stCondLst is None

    def it_rejects_negative_int_delay(self):
        cTn = element("p:cTn")
        with pytest.raises(ValueError, match="non-negative"):
            cTn.delay = -1

    def and_rejects_non_int_non_indefinite_delay(self):
        cTn = element("p:cTn")
        with pytest.raises(ValueError, match="non-negative"):
            cTn.delay = "soon"

    # -- Issue #256: preset descriptors ----------------------------------

    def it_reads_its_presetID_attribute(self):
        cTn = element('p:cTn{presetID=10}')
        assert cTn.presetID == 10

    def and_presetID_is_None_when_absent(self):
        cTn = element("p:cTn")
        assert cTn.presetID is None

    def it_reads_its_presetClass_attribute(self):
        cTn = element('p:cTn{presetClass=entr}')
        assert cTn.presetClass == "entr"

    def it_reads_its_presetSubtype_attribute(self):
        cTn = element('p:cTn{presetSubtype=8}')
        assert cTn.presetSubtype == 8


class DescribeCT_TLTimeCondition(object):
    """Unit-test suite for `pptx.oxml.timing.CT_TLTimeCondition`."""

    def it_is_used_for_p_cond(self):
        cond = element("p:cond")
        assert isinstance(cond, CT_TLTimeCondition)

    def it_reads_its_delay_attribute_as_int(self):
        cond = element("p:cond{delay=750}")
        assert cond.delay == 750

    def and_evt_is_None_when_absent(self):
        cond = element("p:cond")
        assert cond.evt is None

    def and_reads_indefinite_delay_as_string(self):
        cond = element("p:cond{delay=indefinite}")
        assert cond.delay == "indefinite"

    def and_delay_is_None_when_absent(self):
        cond = element("p:cond")
        assert cond.delay is None

    def it_writes_its_delay_attribute(self):
        cond = element("p:cond")
        cond.delay = 250
        assert cond.delay == 250

    def it_reads_its_evt_attribute(self):
        cond = element("p:cond{evt=onClick}")
        assert cond.evt == "onClick"


class DescribeCT_TLTimeConditionList(object):
    """Unit-test suite for `pptx.oxml.timing.CT_TLTimeConditionList`."""

    def it_is_used_for_the_p_stCondLst_element(self):
        stCondLst = element("p:stCondLst")
        assert isinstance(stCondLst, CT_TLTimeConditionList)

    def it_is_used_for_the_p_endCondLst_element(self):
        endCondLst = element("p:endCondLst")
        assert isinstance(endCondLst, CT_TLTimeConditionList)

    def it_can_add_a_blank_cond_child(self):
        stCondLst = element("p:stCondLst")
        cond = stCondLst.add_cond()
        assert isinstance(cond, CT_TLTimeCondition)
        assert cond.delay is None
        assert cond.getparent() is stCondLst


class DescribeCT_TLShapeTargetElement(object):
    """Unit-test suite for `pptx.oxml.timing.CT_TLShapeTargetElement` (issue #256)."""

    def it_is_used_for_the_p_spTgt_element(self):
        spTgt = element("p:spTgt")
        assert isinstance(spTgt, CT_TLShapeTargetElement)

    def it_reads_its_spid_attribute(self):
        spTgt = element('p:spTgt{spid=42}')
        assert spTgt.spid == 42

    def and_spid_is_None_when_absent(self):
        spTgt = element("p:spTgt")
        assert spTgt.spid is None


class Describe_iter_main_sequence_effects(object):
    """Unit-test suite for `pptx.oxml.timing.iter_main_sequence_effects`."""

    def it_yields_nothing_for_a_None_timing(self):
        assert list(iter_main_sequence_effects(None)) == []

    def it_yields_nothing_when_there_is_no_mainSeq(self):
        timing = element("p:timing/p:tnLst/p:par/p:cTn{id=1,nodeType=tmRoot}")
        assert list(iter_main_sequence_effects(timing)) == []

    def it_yields_nothing_for_a_mainSeq_with_no_preset_par(self):
        # -- mainSeq present but with no effect-level `p:par` children --
        timing = element(
            "p:timing/p:tnLst/p:par/(p:cTn{id=1,nodeType=tmRoot}/p:childTnLst/"
            "p:seq/p:cTn{id=2,nodeType=mainSeq})"
        )
        assert list(iter_main_sequence_effects(timing)) == []

    def it_yields_each_effect_par_in_document_order(self, effect_timing):
        effects = list(iter_main_sequence_effects(effect_timing))
        assert len(effects) == 2
        # -- first effect is on shape 3, second on shape 4 --
        assert first_spTgt_spid(effects[0]) == 3
        assert first_spTgt_spid(effects[1]) == 4

    def it_skips_wrapping_par_nodes_without_presetClass(self, effect_timing):
        # -- tree has intermediate (non-effect) `p:par` wrappers; they
        # -- must NOT be yielded.
        for effect in iter_main_sequence_effects(effect_timing):
            cTn = effect[0]  # first child is p:cTn by schema
            assert "presetClass" in cTn.attrib

    # -- fixtures ---------------------------------------------------------

    @pytest.fixture
    def effect_timing(self):
        """p:timing subtree with two effect-level `p:par` nodes in mainSeq."""
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        timing_xml = (
            "<p:timing %s>"
            "  <p:tnLst>"
            '    <p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst>'
            '      <p:seq><p:cTn id="2" nodeType="mainSeq"><p:childTnLst>'
            # -- click-group par --
            '        <p:par><p:cTn id="3"><p:childTnLst>'
            # -- step par --
            '          <p:par><p:cTn id="4"><p:childTnLst>'
            # -- effect par #1 --
            '            <p:par><p:cTn id="5" presetID="1" presetClass="entr">'
            "              <p:childTnLst>"
            "                <p:set><p:cBhvr>"
            '                  <p:cTn id="6"/>'
            '                  <p:tgtEl><p:spTgt spid="3"/></p:tgtEl>'
            "                </p:cBhvr></p:set>"
            "              </p:childTnLst>"
            "            </p:cTn></p:par>"
            # -- effect par #2 --
            '            <p:par><p:cTn id="7" presetID="10" presetClass="exit">'
            "              <p:childTnLst>"
            "                <p:set><p:cBhvr>"
            '                  <p:cTn id="8"/>'
            '                  <p:tgtEl><p:spTgt spid="4"/></p:tgtEl>'
            "                </p:cBhvr></p:set>"
            "              </p:childTnLst>"
            "            </p:cTn></p:par>"
            "          </p:childTnLst></p:cTn></p:par>"
            "        </p:childTnLst></p:cTn></p:par>"
            "      </p:childTnLst></p:cTn></p:seq>"
            "    </p:childTnLst></p:cTn></p:par>"
            "  </p:tnLst>"
            "</p:timing>"
        ) % nsdecls("p")
        return parse_xml(timing_xml)


class Describe_first_spTgt_spid(object):
    """Unit-test suite for `pptx.oxml.timing.first_spTgt_spid` (issue #256)."""

    def it_returns_None_when_no_spTgt_descendant(self):
        par = element("p:par/p:cTn")
        assert first_spTgt_spid(par) is None

    def it_returns_spid_of_first_spTgt_descendant(self):
        par = element(
            "p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:tgtEl/p:spTgt{spid=7}"
        )
        assert first_spTgt_spid(par) == 7


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
