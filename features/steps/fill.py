# encoding: utf-8

"""Gherkin step implementations for FillFormat-related features."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from behave import given, then, when

from pptx import Presentation
from pptx.enum.dml import MSO_FILL, MSO_PATTERN

from helpers import test_pptx


# given ====================================================

@given('a FillFormat object as fill')
def given_a_FillFormat_object_as_fill(context):
    fill = Presentation(test_pptx('dml-fill')).slides[0].shapes[0].fill
    context.fill = fill


@given('a FillFormat object as fill having {pattern} fill')
def given_a_FillFormat_object_as_fill_having_pattern(context, pattern):
    shape_idx = {
        'no pattern':        0,
        'MSO_PATTERN.DIVOT': 1,
        'MSO_PATTERN.WAVE':  2,
    }[pattern]
    slide = Presentation(test_pptx('dml-fill')).slides[1]
    fill = slide.shapes[shape_idx].fill
    context.fill = fill


# when =====================================================

@when("I assign {value} to fill.pattern")
def when_I_assign_value_to_fill_pattern(context, value):
    pattern = {
        'None':              None,
        'MSO_PATTERN.CROSS': MSO_PATTERN.CROSS,
        'MSO_PATTERN.DIVOT': MSO_PATTERN.DIVOT,
        'MSO_PATTERN.WAVE':  MSO_PATTERN.WAVE,
    }[value]
    context.fill.pattern = pattern


@when("I call fill.background()")
def when_I_call_fill_background(context):
    context.fill.background()


@when("I call fill.patterned()")
def when_I_call_fill_patterned(context):
    context.fill.patterned()


@when("I call fill.solid()")
def when_I_call_fill_solid(context):
    context.fill.solid()


# then =====================================================

@then('fill.back_color is a ColorFormat object')
def then_fill_back_color_is_a_ColorFormat_object(context):
    class_name = context.fill.back_color.__class__.__name__
    expected_value = 'ColorFormat'
    assert class_name == expected_value, (
        'expected \'%s\', got \'%s\'' % (expected_value, class_name)
    )


@then('fill.fore_color is a ColorFormat object')
def then_fill_fore_color_is_a_ColorFormat_object(context):
    class_name = context.fill.fore_color.__class__.__name__
    expected_value = 'ColorFormat'
    assert class_name == expected_value, (
        'expected \'%s\', got \'%s\'' % (expected_value, class_name)
    )


@then('fill.pattern is {value}')
def then_fill_pattern_is_value(context, value):
    fill_pattern = context.fill.pattern
    expected_value = {
        'None':              None,
        'MSO_PATTERN.CROSS': MSO_PATTERN.CROSS,
        'MSO_PATTERN.DIVOT': MSO_PATTERN.DIVOT,
        'MSO_PATTERN.WAVE':  MSO_PATTERN.WAVE,
    }[value]
    assert fill_pattern == expected_value, (
        'expected fill pattern %s, got %s' % (expected_value, fill_pattern)
    )


@then('fill.type is MSO_FILL.{member_name}')
def then_fill_type_is_MSO_FILL_member_name(context, member_name):
    fill_type = context.fill.type
    expected_value = getattr(MSO_FILL, member_name)
    assert fill_type == expected_value, (
        'expected fill type %s, got %s' % (expected_value, fill_type)
    )
