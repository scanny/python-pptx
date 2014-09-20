# encoding: utf-8

"""
Step implementations for run property (font)-related features
"""

from __future__ import absolute_import

from behave import given, then

from pptx import Presentation

from .helpers import test_pptx


# given ===================================================

@given('a font')
def given_a_font(context):
    prs = Presentation(test_pptx('txt-font-props'))
    slide = prs.slides[1]
    textbox = slide.shapes[0]
    run = textbox.text_frame.paragraphs[0].runs[0]
    context.font = run.font


@given('a font having {applied_size}')
def given_a_font_having_applied_size(context, applied_size):
    shape_idx = {
        'no directly applied size':        0,
        'a directly applied size of 42pt': 1,
    }[applied_size]
    prs = Presentation(test_pptx('txt-font-props'))
    slide = prs.slides[1]
    textbox = slide.shapes[shape_idx]
    run = textbox.text_frame.paragraphs[0].runs[0]
    context.font = run.font


# then ===================================================

@then('the font size property value is {value_str}')
def then_the_font_size_property_value_is_value(context, value_str):
    expected_value = {
        '42.0 points': 42.0,
        'None':        None,
    }[value_str]
    font = context.font
    value = font.size if font.size is None else font.size.pt
    assert value == expected_value, (
        'expected %s, got %s' % (expected_value, value)
    )
