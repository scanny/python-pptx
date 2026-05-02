"""Gherkin step implementations for slide-transition round-trip (Foundation F8)."""

from __future__ import annotations

import io

from behave import then, when

from pptx import Presentation
from pptx.enum.transition import PP_TRANSITION_TYPE


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
