"""Unit-test suite for `pptx.animation` module (issue #102 MVP)."""

from __future__ import annotations

import pytest

from pptx.animation import (
    AnimationEffect,
    _PRESET_MAP,
    _ensure_main_sequence,
    _find_effect_par_for_spid,
    _next_cTn_id,
    _preset_to_type,
    _remove_effect_par,
    _set_animation,
)
from pptx.enum.animation import MSO_ANIMATION_TRIGGER, MSO_ANIMATION_TYPE

from .unitutil.cxml import element


class Describe_preset_to_type(object):
    """Unit-test suite for the `_preset_to_type` helper."""

    @pytest.mark.parametrize(
        ("preset_id", "preset_class", "expected"),
        [
            (1, "entr", MSO_ANIMATION_TYPE.APPEAR),
            (10, "entr", MSO_ANIMATION_TYPE.FADE_IN),
            (2, "entr", MSO_ANIMATION_TYPE.FLY_IN),
            (1, "emph", MSO_ANIMATION_TYPE.PULSE),
            (10, "exit", MSO_ANIMATION_TYPE.FADE_OUT),
        ],
    )
    def it_maps_known_presets_to_enum_members(self, preset_id, preset_class, expected):
        assert _preset_to_type(preset_id, preset_class) is expected

    def it_returns_NONE_for_unknown_presets(self):
        assert _preset_to_type(999, "entr") is MSO_ANIMATION_TYPE.NONE

    def it_returns_NONE_when_preset_id_is_missing(self):
        assert _preset_to_type(None, "entr") is MSO_ANIMATION_TYPE.NONE

    def it_returns_NONE_when_preset_class_is_missing(self):
        assert _preset_to_type(10, None) is MSO_ANIMATION_TYPE.NONE


class DescribeAnimationEffect(object):
    """Unit-test suite for the AnimationEffect proxy class."""

    def it_reports_its_shape_id_via_spTgt(self):
        par = _par_with_preset(10, "entr", "clickEffect", delay=500, spid=42)
        effect = AnimationEffect(par)
        assert effect.shape_id == 42

    def it_reports_NONE_for_shape_id_when_no_spTgt(self):
        par = _par_with_preset(10, "entr", "clickEffect", delay=0, spid=None)
        effect = AnimationEffect(par)
        assert effect.shape_id is None

    @pytest.mark.parametrize(
        ("preset_id", "preset_class", "expected"),
        [
            (1, "entr", MSO_ANIMATION_TYPE.APPEAR),
            (10, "entr", MSO_ANIMATION_TYPE.FADE_IN),
            (2, "entr", MSO_ANIMATION_TYPE.FLY_IN),
            (1, "emph", MSO_ANIMATION_TYPE.PULSE),
            (10, "exit", MSO_ANIMATION_TYPE.FADE_OUT),
        ],
    )
    def it_reports_its_type_from_preset_attrs(self, preset_id, preset_class, expected):
        par = _par_with_preset(preset_id, preset_class, "clickEffect", delay=0, spid=1)
        effect = AnimationEffect(par)
        assert effect.type is expected

    def it_reports_type_NONE_for_unknown_preset(self):
        par = _par_with_preset(999, "path", "clickEffect", delay=0, spid=1)
        effect = AnimationEffect(par)
        assert effect.type is MSO_ANIMATION_TYPE.NONE

    def it_reports_type_NONE_when_preset_attrs_absent(self):
        par = _par_without_preset()
        effect = AnimationEffect(par)
        assert effect.type is MSO_ANIMATION_TYPE.NONE

    def it_reports_ON_CLICK_trigger_for_clickEffect_nodeType(self):
        par = _par_with_preset(10, "entr", "clickEffect", delay=0, spid=1)
        effect = AnimationEffect(par)
        assert effect.trigger is MSO_ANIMATION_TRIGGER.ON_CLICK

    def it_reports_AFTER_PREVIOUS_trigger_for_afterEffect_nodeType(self):
        par = _par_with_preset(10, "entr", "afterEffect", delay=0, spid=1)
        effect = AnimationEffect(par)
        assert effect.trigger is MSO_ANIMATION_TRIGGER.AFTER_PREVIOUS

    def it_falls_back_to_ON_CLICK_for_unknown_nodeType(self):
        par = _par_with_preset(10, "entr", "withEffect", delay=0, spid=1)
        effect = AnimationEffect(par)
        assert effect.trigger is MSO_ANIMATION_TRIGGER.ON_CLICK

    def it_reports_its_delay(self):
        par = _par_with_preset(10, "entr", "clickEffect", delay=1500, spid=1)
        effect = AnimationEffect(par)
        assert effect.delay == 1500

    def it_reports_zero_delay_when_no_stCondLst(self):
        par = _par_without_preset()
        effect = AnimationEffect(par)
        assert effect.delay == 0


class Describe_next_cTn_id(object):
    """Unit-test suite for `_next_cTn_id` helper."""

    def it_returns_1_when_no_timing_element(self):
        sld = element("p:sld")
        assert _next_cTn_id(sld) == 1

    def it_returns_one_greater_than_max_existing_id(self):
        sld = _sld_with_timing(
            '<p:timing %s>'
            '  <p:tnLst>'
            '    <p:par><p:cTn id="1"/></p:par>'
            '    <p:par><p:cTn id="5"/></p:par>'
            '    <p:par><p:cTn id="3"/></p:par>'
            '  </p:tnLst>'
            '</p:timing>'
        )
        assert _next_cTn_id(sld) == 6


class Describe_ensure_main_sequence(object):
    """Unit-test suite for `_ensure_main_sequence` helper."""

    def it_creates_mainSeq_from_scratch_on_bare_slide(self):
        sld = element("p:sld")
        mainSeq_cTn = _ensure_main_sequence(sld)
        assert mainSeq_cTn.nodeType == "mainSeq"
        # -- and it lives inside timing/tnLst/par[tmRoot]/cTn/childTnLst/seq/cTn
        matches = sld.xpath(
            "./p:timing/p:tnLst/p:par"
            "[p:cTn[@nodeType='tmRoot']]"
            "/p:cTn/p:childTnLst/p:seq"
            "[p:cTn[@nodeType='mainSeq']]"
        )
        assert len(matches) == 1

    def it_returns_existing_mainSeq_when_present(self):
        sld = _sld_with_full_mainSeq()
        mainSeq_cTn_1 = _ensure_main_sequence(sld)
        mainSeq_cTn_2 = _ensure_main_sequence(sld)
        assert mainSeq_cTn_1 is mainSeq_cTn_2
        assert mainSeq_cTn_1.nodeType == "mainSeq"


class Describe_find_effect_par_for_spid(object):
    """Unit-test suite for `_find_effect_par_for_spid`."""

    def it_returns_None_when_no_timing_element(self):
        sld = element("p:sld")
        assert _find_effect_par_for_spid(sld, 5) is None

    def it_returns_None_when_no_matching_spTgt(self):
        sld = _sld_with_full_mainSeq()
        assert _find_effect_par_for_spid(sld, 999) is None


class Describe_set_animation(object):
    """Unit-test suite for the `_set_animation` helper."""

    def it_rejects_NONE_effect_type(self):
        sld = element("p:sld")
        with pytest.raises(ValueError, match="MVP presets"):
            _set_animation(
                sld, 1, MSO_ANIMATION_TYPE.NONE, MSO_ANIMATION_TRIGGER.ON_CLICK, 0
            )

    def it_rejects_negative_delay(self):
        sld = element("p:sld")
        with pytest.raises(ValueError, match="non-negative"):
            _set_animation(
                sld, 1, MSO_ANIMATION_TYPE.FADE_IN, MSO_ANIMATION_TRIGGER.ON_CLICK, -1
            )

    def it_writes_fade_in_preset_attrs_correctly(self):
        sld = element("p:sld")
        effect = _set_animation(
            sld, 7, MSO_ANIMATION_TYPE.FADE_IN, MSO_ANIMATION_TRIGGER.ON_CLICK, 0
        )
        assert effect.type is MSO_ANIMATION_TYPE.FADE_IN
        assert effect.shape_id == 7

    def it_replaces_a_prior_effect_on_the_same_shape(self):
        sld = element("p:sld")
        _set_animation(
            sld, 7, MSO_ANIMATION_TYPE.FADE_IN, MSO_ANIMATION_TRIGGER.ON_CLICK, 0
        )
        _set_animation(
            sld, 7, MSO_ANIMATION_TYPE.PULSE, MSO_ANIMATION_TRIGGER.AFTER_PREVIOUS, 100
        )
        matches = sld.xpath(".//p:spTgt[@spid='7']")
        assert len(matches) == 1
        # -- now the single matching effect should be PULSE --
        par = _find_effect_par_for_spid(sld, 7)
        effect = AnimationEffect(par)
        assert effect.type is MSO_ANIMATION_TYPE.PULSE
        assert effect.trigger is MSO_ANIMATION_TRIGGER.AFTER_PREVIOUS
        assert effect.delay == 100

    def it_can_add_multiple_effects_for_different_shapes(self):
        sld = element("p:sld")
        _set_animation(
            sld, 7, MSO_ANIMATION_TYPE.FADE_IN, MSO_ANIMATION_TRIGGER.ON_CLICK, 0
        )
        _set_animation(
            sld, 8, MSO_ANIMATION_TYPE.FLY_IN, MSO_ANIMATION_TRIGGER.ON_CLICK, 0
        )
        assert _find_effect_par_for_spid(sld, 7) is not None
        assert _find_effect_par_for_spid(sld, 8) is not None

    def it_writes_APPEAR_as_a_p_set_behavior(self):
        sld = element("p:sld")
        _set_animation(
            sld, 1, MSO_ANIMATION_TYPE.APPEAR, MSO_ANIMATION_TRIGGER.ON_CLICK, 0
        )
        assert len(sld.xpath(".//p:set")) == 1
        assert len(sld.xpath(".//p:animEffect")) == 0

    def it_writes_FADE_IN_as_p_animEffect_with_transition_in(self):
        sld = element("p:sld")
        _set_animation(
            sld, 1, MSO_ANIMATION_TYPE.FADE_IN, MSO_ANIMATION_TRIGGER.ON_CLICK, 0
        )
        animEffects = sld.xpath(".//p:animEffect")
        assert len(animEffects) == 1
        assert animEffects[0].get("transition") == "in"
        assert animEffects[0].get("filter") == "fade"

    def it_writes_FLY_IN_as_p_animEffect_with_slide_filter(self):
        sld = element("p:sld")
        _set_animation(
            sld, 1, MSO_ANIMATION_TYPE.FLY_IN, MSO_ANIMATION_TRIGGER.ON_CLICK, 0
        )
        animEffects = sld.xpath(".//p:animEffect")
        assert len(animEffects) == 1
        assert animEffects[0].get("filter") == "slide(fromBottom)"

    def it_writes_FADE_OUT_as_p_animEffect_with_transition_out(self):
        sld = element("p:sld")
        _set_animation(
            sld, 1, MSO_ANIMATION_TYPE.FADE_OUT, MSO_ANIMATION_TRIGGER.ON_CLICK, 0
        )
        animEffects = sld.xpath(".//p:animEffect")
        assert len(animEffects) == 1
        assert animEffects[0].get("transition") == "out"


class Describe_remove_effect_par(object):
    """Unit-test suite for `_remove_effect_par`."""

    def it_removes_par_from_parent(self):
        sld = element("p:sld")
        _set_animation(
            sld, 7, MSO_ANIMATION_TYPE.FADE_IN, MSO_ANIMATION_TRIGGER.ON_CLICK, 0
        )
        par = _find_effect_par_for_spid(sld, 7)
        assert par is not None
        _remove_effect_par(sld, par)
        assert _find_effect_par_for_spid(sld, 7) is None


# -- helpers ----------------------------------------------------------------


def _par_with_preset(preset_id, preset_class, node_type, delay, spid):
    """Return a `p:par` element configured per arguments.

    Minimal: one cTn child with preset attrs and optional stCondLst +
    inner spTgt.
    """
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls

    spTgt_xml = (
        '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>' % spid if spid is not None else ""
    )
    behavior_xml = (
        "<p:animEffect>"
        "<p:cBhvr>"
        '<p:cTn id="99"/>'
        "%s"
        "</p:cBhvr>"
        "</p:animEffect>" % spTgt_xml
    )
    xml_str = (
        "<p:par %s>"
        '<p:cTn id="3" presetID="%d" presetClass="%s" nodeType="%s">'
        '<p:stCondLst><p:cond delay="%d"/></p:stCondLst>'
        "<p:childTnLst>%s</p:childTnLst>"
        "</p:cTn>"
        "</p:par>"
        % (nsdecls("p"), preset_id, preset_class, node_type, delay, behavior_xml)
    )
    return parse_xml(xml_str)


def _par_without_preset():
    """Return a `p:par` with no preset attributes or stCondLst."""
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls

    xml_str = '<p:par %s><p:cTn id="3"/></p:par>' % nsdecls("p")
    return parse_xml(xml_str)


def _sld_with_timing(xml_tpl):
    """Return a CT_Slide with the given `<p:timing>` inserted."""
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls

    sld_xml = (
        "<p:sld %s>"
        "<p:cSld><p:spTree>"
        "<p:nvGrpSpPr>"
        '<p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
        "</p:nvGrpSpPr>"
        "<p:grpSpPr/>"
        "</p:spTree></p:cSld>"
        "%s"
        "</p:sld>" % (nsdecls("a", "p", "r"), xml_tpl % nsdecls("p"))
    )
    return parse_xml(sld_xml)


def _sld_with_full_mainSeq():
    """Return a CT_Slide with a complete main-sequence skeleton."""
    tpl = (
        "<p:timing %s>"
        "<p:tnLst>"
        "<p:par>"
        '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        "<p:childTnLst>"
        '<p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        "<p:childTnLst/>"
        "</p:cTn>"
        "</p:seq>"
        "</p:childTnLst>"
        "</p:cTn>"
        "</p:par>"
        "</p:tnLst>"
        "</p:timing>"
    )
    return _sld_with_timing(tpl)
