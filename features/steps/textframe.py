# encoding: utf-8

"""
Step implementations for textframe-related features
"""

from __future__ import absolute_import

from behave import given, then, when

from pptx import Presentation
from pptx.util import Inches


# given ===================================================

@given('a textframe')
def step_given_a_textframe(context):
    context.prs = Presentation()
    blank_slide_layout = context.prs.slide_layouts[6]
    slide = context.prs.slides.add_slide(blank_slide_layout)
    length = Inches(2.00)
    textbox = slide.shapes.add_textbox(length, length, length, length)
    context.textframe = textbox.textframe


# when ====================================================

@when("I set the textframe word wrap {setting}")
def when_set_textframe_word_wrap(context, setting):
    bool_val = {'on': True, 'off': False, 'to None': None}
    context.textframe.word_wrap = bool_val[setting]


# then ====================================================

@then('the textframe\'s {side} margin is {inches}"')
def then_textframe_margin_is_value(context, side, inches):
    textframe = context.prs.slides[0].shapes[0].textframe
    emu = Inches(float(inches))
    if side == 'left':
        assert textframe.margin_left == emu
    elif side == 'top':
        assert textframe.margin_top == emu
    elif side == 'right':
        assert textframe.margin_right == emu
    elif side == 'bottom':
        assert textframe.margin_bottom == emu


@then('the textframe word wrap is set {setting}')
def then_textframe_word_wrap_is_setting(context, setting):
    expected_value = {
        'on': True, 'off': False, 'to None': None
    }[setting]
    textframe = context.prs.slides[0].shapes[0].textframe
    assert textframe.word_wrap is expected_value
