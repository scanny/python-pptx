"""Gherkin step implementations for slide-transition round-trip (Foundation F8)."""

from __future__ import annotations

import io

from behave import given, then, when
from pptx import Presentation
from pptx.enum.transition import PP_TRANSITION_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn


# when ============================================================


@when("I set slide.transition.type to {name}")
def when_set_slide_transition_type(context, name):
    context.slide.transition.type = getattr(PP_TRANSITION_TYPE, name)


@when("I set slide.transition.duration to {ms:d}")
def when_set_slide_transition_duration(context, ms):
    context.slide.transition.duration = ms


@when("I set slide.transition.advance_on_click to {value}")
def when_set_slide_transition_advance_on_click(context, value):
    context.slide.transition.advance_on_click = value == "True"


@when("I set slide.transition.advance_after_time to {ms:d}")
def when_set_slide_transition_advance_after_time(context, ms):
    context.slide.transition.advance_after_time = ms


@when("I save and reopen the presentation")
def when_save_and_reopen_presentation(context):
    buf = io.BytesIO()
    context.prs.save(buf)
    buf.seek(0)
    context.prs = Presentation(buf)
    context.slide = context.prs.slides[0]


# then ============================================================


@then("slide.transition.type is PP_TRANSITION_TYPE.{name}")
def then_slide_transition_type_is(context, name):
    expected = getattr(PP_TRANSITION_TYPE, name)
    actual = context.slide.transition.type
    assert actual is expected, "slide.transition.type is %s (expected %s)" % (actual, expected)


@then("slide.transition.duration is {ms:d}")
def then_slide_transition_duration_is(context, ms):
    actual = context.slide.transition.duration
    assert actual == ms, "slide.transition.duration is %s" % actual


@then("slide.transition.advance_on_click is {value}")
def then_slide_transition_advance_on_click_is(context, value):
    expected = value == "True"
    actual = context.slide.transition.advance_on_click
    assert actual is expected, "slide.transition.advance_on_click is %s" % actual


@then("slide.transition.advance_after_time is {value}")
def then_slide_transition_advance_after_time_is(context, value):
    if value == "None":
        expected = None
    else:
        expected = int(value)
    actual = context.slide.transition.advance_after_time
    assert actual == expected, "slide.transition.advance_after_time is %s" % actual


@then("slide.has_animations is {value}")
def then_slide_has_animations_is(context, value):
    expected = value == "True"
    actual = context.slide.has_animations
    assert actual is expected, "slide.has_animations is %s" % actual


@then("slide.timing_xml is None")
def then_slide_timing_xml_is_None(context):
    assert context.slide.timing_xml is None, "slide.timing_xml is %s" % context.slide.timing_xml


# given ===========================================================
# -- Issue #256: animation_sequence scenarios --


@given("a slide with two authored entrance effects")
def given_a_slide_with_two_authored_entrance_effects(context):
    """Build a blank slide and stitch a two-effect main sequence into its XML.

    This gives the #256 scenario a realistic ``p:timing/p:tnLst/p:par/.../
    p:seq`` mainSeq with two effect-level ``p:par`` nodes — one ``entr``
    on shape 3 and one ``exit`` on shape 4 — without requiring a
    structured authoring API (that is downstream #102 / #1106).
    """
    context.prs = Presentation()
    slide_layout = context.prs.slide_layouts[6]  # blank layout
    context.slide = context.prs.slides.add_slide(slide_layout)
    sld = context.slide._element
    # -- remove any existing timing (unlikely on a blank slide) --
    existing = sld.find(qn("p:timing"))
    if existing is not None:
        sld.remove(existing)
    timing_xml = (
        "<p:timing %s>"
        "  <p:tnLst>"
        '    <p:par><p:cTn id="1" dur="indefinite" restart="never" '
        'nodeType="tmRoot"><p:childTnLst>'
        '      <p:seq concurrent="1" nextAc="seek">'
        '        <p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        "          <p:childTnLst>"
        '            <p:par><p:cTn id="3" fill="hold"><p:childTnLst>'
        '              <p:par><p:cTn id="4" fill="hold"><p:childTnLst>'
        '                <p:par><p:cTn id="5" presetID="1" presetClass="entr" '
        'presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">'
        "                  <p:childTnLst><p:set><p:cBhvr>"
        '                    <p:cTn id="6" dur="1" fill="hold"/>'
        '                    <p:tgtEl><p:spTgt spid="3"/></p:tgtEl>'
        "                  </p:cBhvr></p:set></p:childTnLst>"
        "                </p:cTn></p:par>"
        '                <p:par><p:cTn id="7" presetID="10" presetClass="exit" '
        'presetSubtype="0" fill="hold" grpId="0" nodeType="afterEffect">'
        "                  <p:childTnLst><p:set><p:cBhvr>"
        '                    <p:cTn id="8" dur="1" fill="hold"/>'
        '                    <p:tgtEl><p:spTgt spid="4"/></p:tgtEl>'
        "                  </p:cBhvr></p:set></p:childTnLst>"
        "                </p:cTn></p:par>"
        "              </p:childTnLst></p:cTn></p:par>"
        "            </p:childTnLst></p:cTn></p:par>"
        "          </p:childTnLst>"
        "        </p:cTn>"
        "      </p:seq>"
        "    </p:childTnLst></p:cTn></p:par>"
        "  </p:tnLst>"
        "</p:timing>"
    ) % nsdecls("p")
    sld.append(parse_xml(timing_xml))


# then ============================================================
# -- Issue #256: animation_sequence assertions --


@then("slide.animation_sequence is an empty tuple")
def then_slide_animation_sequence_is_empty(context):
    seq = context.slide.animation_sequence
    assert seq == (), "slide.animation_sequence is %r" % (seq,)


@then("len(slide.animation_sequence) is {count:d}")
def then_len_slide_animation_sequence_is(context, count):
    actual = len(context.slide.animation_sequence)
    assert actual == count, "len(slide.animation_sequence) is %d" % actual


@then("slide.animation_sequence[{idx:d}].shape_id is {sid:d}")
def then_animation_sequence_idx_shape_id_is(context, idx, sid):
    actual = context.slide.animation_sequence[idx].shape_id
    assert actual == sid, (
        "slide.animation_sequence[%d].shape_id is %s" % (idx, actual)
    )


@then("slide.animation_sequence[{idx:d}].preset_class is '{value}'")
def then_animation_sequence_idx_preset_class_is(context, idx, value):
    actual = context.slide.animation_sequence[idx].preset_class
    assert actual == value, (
        "slide.animation_sequence[%d].preset_class is %r" % (idx, actual)
    )


@then("slide.animation_sequence[{idx:d}].preset_id is {value:d}")
def then_animation_sequence_idx_preset_id_is(context, idx, value):
    actual = context.slide.animation_sequence[idx].preset_id
    assert actual == value, (
        "slide.animation_sequence[%d].preset_id is %s" % (idx, actual)
    )
