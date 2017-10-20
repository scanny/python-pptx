# encoding: utf-8

"""Gherkin step implementations for FillFormat-related features."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from behave import given, then, when

from pptx import Presentation
from pptx.enum.dml import MSO_FILL

from helpers import test_pptx


# given ====================================================

@given('a FillFormat object as fill')
def given_a_FillFormat_object_as_fill(context):
    fill = Presentation(test_pptx('dml-fill')).slides[0].shapes[0].fill
    context.fill = fill


# when =====================================================

@when("I call fill.background()")
def when_I_call_fill_background(context):
    context.fill.background()


@when("I call fill.solid()")
def when_I_call_fill_solid(context):
    context.fill.solid()


# then =====================================================

@then('fill.fore_color is a ColorFormat object')
def then_fill_fore_color_is_a_ColorFormat_object(context):
    class_name = context.fill.fore_color.__class__.__name__
    expected_value = 'ColorFormat'
    assert class_name == expected_value, (
        'expected \'%s\', got \'%s\'' % (expected_value, class_name)
    )


@then('fill.type is MSO_FILL.{member_name}')
def then_fill_type_is_MSO_FILL_member_name(context, member_name):
    fill_type = context.fill.type
    expected_value = getattr(MSO_FILL, member_name)
    assert fill_type == expected_value, (
        'expected fill type %s, got %s' % (expected_value, fill_type)
    )
