"""Unit-test suite for `pptx.oxml.animation` module (issue #102 MVP)."""

from __future__ import annotations

import pytest

from pptx.oxml.animation import (
    CT_TLAnimateBehavior,
    CT_TLAnimateEffectBehavior,
    CT_TLCommonBehaviorData,
    CT_TLSetBehavior,
    CT_TLShapeTargetElement,
    CT_TLTimeCondition,
    CT_TLTimeConditionList,
    CT_TLTimeTargetElement,
    ST_TLTriggerEvent,
)

from ..unitutil.cxml import element


class DescribeST_TLTriggerEvent(object):
    """Unit-test suite for the ST_TLTriggerEvent simple type."""

    def it_passes_valid_tokens_through_convert_from_xml(self):
        assert ST_TLTriggerEvent.convert_from_xml("onClick") == "onClick"

    def it_passes_valid_tokens_through_convert_to_xml(self):
        assert ST_TLTriggerEvent.convert_to_xml("onPrev") == "onPrev"

    @pytest.mark.parametrize(
        "value", ["onClick", "onNext", "onPrev", "onMouseOver", "onStopAudio"]
    )
    def it_accepts_valid_tokens_on_validate(self, value):
        ST_TLTriggerEvent.validate(value)

    def it_rejects_unknown_tokens_on_validate(self):
        with pytest.raises(ValueError, match="ST_TLTriggerEvent"):
            ST_TLTriggerEvent.validate("onSomethingElse")

    def it_rejects_non_string_on_validate(self):
        with pytest.raises(ValueError):
            ST_TLTriggerEvent.validate(42)


class DescribeCT_TLShapeTargetElement(object):
    """Unit-test suite for `CT_TLShapeTargetElement`."""

    def it_is_the_class_used_for_the_p_spTgt_element(self):
        spTgt = element("p:spTgt{spid=5}")
        assert isinstance(spTgt, CT_TLShapeTargetElement)

    def it_reads_its_spid_attribute_as_int(self):
        spTgt = element("p:spTgt{spid=12}")
        assert spTgt.spid == 12

    def it_can_set_its_spid_attribute(self):
        spTgt = element("p:spTgt{spid=1}")
        spTgt.spid = 7
        assert spTgt.get("spid") == "7"


class DescribeCT_TLTimeTargetElement(object):
    """Unit-test suite for `CT_TLTimeTargetElement`."""

    def it_is_the_class_used_for_the_p_tgtEl_element(self):
        tgtEl = element("p:tgtEl")
        assert isinstance(tgtEl, CT_TLTimeTargetElement)

    def it_exposes_an_optional_spTgt_child(self):
        tgtEl = element("p:tgtEl/p:spTgt{spid=3}")
        assert tgtEl.spTgt is not None
        assert tgtEl.spTgt.spid == 3

    def and_spTgt_is_None_when_absent(self):
        tgtEl = element("p:tgtEl")
        assert tgtEl.spTgt is None


class DescribeCT_TLTimeCondition(object):
    """Unit-test suite for `CT_TLTimeCondition`."""

    def it_is_the_class_used_for_the_p_cond_element(self):
        cond = element("p:cond")
        assert isinstance(cond, CT_TLTimeCondition)

    def it_reads_its_evt_attribute(self):
        cond = element("p:cond{evt=onClick}")
        assert cond.evt == "onClick"

    def it_reads_its_delay_attribute_as_int(self):
        cond = element("p:cond{delay=500}")
        assert cond.delay == 500

    def and_reads_indefinite_delay_as_string(self):
        cond = element("p:cond{delay=indefinite}")
        assert cond.delay == "indefinite"

    def and_delay_is_None_when_absent(self):
        cond = element("p:cond")
        assert cond.delay is None


class DescribeCT_TLTimeConditionList(object):
    """Unit-test suite for `CT_TLTimeConditionList`."""

    def it_is_used_for_p_stCondLst(self):
        stCondLst = element("p:stCondLst")
        assert isinstance(stCondLst, CT_TLTimeConditionList)

    def it_is_used_for_p_endCondLst(self):
        endCondLst = element("p:endCondLst")
        assert isinstance(endCondLst, CT_TLTimeConditionList)

    def it_is_used_for_p_prevCondLst(self):
        prevCondLst = element("p:prevCondLst")
        assert isinstance(prevCondLst, CT_TLTimeConditionList)

    def it_is_used_for_p_nextCondLst(self):
        nextCondLst = element("p:nextCondLst")
        assert isinstance(nextCondLst, CT_TLTimeConditionList)


class DescribeCT_TLCommonBehaviorData(object):
    """Unit-test suite for `CT_TLCommonBehaviorData`."""

    def it_is_the_class_used_for_the_p_cBhvr_element(self):
        cBhvr = element("p:cBhvr/(p:cTn,p:tgtEl)")
        assert isinstance(cBhvr, CT_TLCommonBehaviorData)

    def it_exposes_its_required_cTn_child(self):
        cBhvr = element("p:cBhvr/(p:cTn{id=4},p:tgtEl)")
        assert cBhvr.cTn.id == 4

    def it_exposes_its_required_tgtEl_child(self):
        cBhvr = element("p:cBhvr/(p:cTn,p:tgtEl/p:spTgt{spid=7})")
        assert cBhvr.tgtEl.spTgt.spid == 7

    def it_reports_spid_via_spid_property(self):
        cBhvr = element("p:cBhvr/(p:cTn,p:tgtEl/p:spTgt{spid=9})")
        assert cBhvr.spid == 9

    def and_spid_is_None_when_target_is_not_a_shape(self):
        cBhvr = element("p:cBhvr/(p:cTn,p:tgtEl)")
        assert cBhvr.spid is None


class DescribeCT_TLSetBehavior(object):
    """Unit-test suite for `CT_TLSetBehavior` (p:set)."""

    def it_is_the_class_used_for_the_p_set_element(self):
        setElm = element("p:set/p:cBhvr/(p:cTn,p:tgtEl)")
        assert isinstance(setElm, CT_TLSetBehavior)

    def it_exposes_its_required_cBhvr_child(self):
        setElm = element("p:set/p:cBhvr/(p:cTn{id=3},p:tgtEl)")
        assert setElm.cBhvr.cTn.id == 3


class DescribeCT_TLAnimateBehavior(object):
    """Unit-test suite for `CT_TLAnimateBehavior` (p:anim)."""

    def it_is_the_class_used_for_the_p_anim_element(self):
        anim = element("p:anim/p:cBhvr/(p:cTn,p:tgtEl)")
        assert isinstance(anim, CT_TLAnimateBehavior)

    def it_exposes_its_required_cBhvr_child(self):
        anim = element("p:anim/p:cBhvr/(p:cTn{id=5},p:tgtEl)")
        assert anim.cBhvr.cTn.id == 5


class DescribeCT_TLAnimateEffectBehavior(object):
    """Unit-test suite for `CT_TLAnimateEffectBehavior` (p:animEffect)."""

    def it_is_the_class_used_for_the_p_animEffect_element(self):
        animEffect = element("p:animEffect/p:cBhvr/(p:cTn,p:tgtEl)")
        assert isinstance(animEffect, CT_TLAnimateEffectBehavior)

    def it_reads_its_transition_attribute(self):
        animEffect = element(
            "p:animEffect{transition=in}/p:cBhvr/(p:cTn,p:tgtEl)"
        )
        assert animEffect.transition == "in"

    def it_reads_its_filter_attribute(self):
        animEffect = element(
            "p:animEffect{filter=fade}/p:cBhvr/(p:cTn,p:tgtEl)"
        )
        assert animEffect.filter_attr == "fade"

    def and_transition_is_None_when_absent(self):
        animEffect = element("p:animEffect/p:cBhvr/(p:cTn,p:tgtEl)")
        assert animEffect.transition is None


class DescribeCT_TLCommonTimeNodeData_preset_attrs(object):
    """Unit-test suite for the new preset* descriptors on CT_TLCommonTimeNodeData.

    The class lives in :mod:`pptx.oxml.timing` but gains presetID /
    presetClass / presetSubtype in the issue #102 MVP.
    """

    def it_reads_presetID_as_int(self):
        cTn = element("p:cTn{presetID=10}")
        assert cTn.presetID == 10

    def it_reads_presetClass(self):
        cTn = element("p:cTn{presetClass=entr}")
        assert cTn.presetClass == "entr"

    def it_reads_presetSubtype_as_int(self):
        cTn = element("p:cTn{presetSubtype=4}")
        assert cTn.presetSubtype == 4

    def and_presetID_is_None_when_absent(self):
        cTn = element("p:cTn")
        assert cTn.presetID is None

    def it_exposes_stCondLst_as_ZeroOrOne(self):
        cTn = element("p:cTn/p:stCondLst/p:cond{delay=500}")
        assert cTn.stCondLst is not None

    def and_stCondLst_is_None_when_absent(self):
        cTn = element("p:cTn")
        assert cTn.stCondLst is None

    def it_exposes_childTnLst_as_ZeroOrOne(self):
        cTn = element("p:cTn/p:childTnLst")
        assert cTn.childTnLst is not None
