# encoding: utf-8

"""
Gherkin step implementations for text-related features.
"""

from __future__ import absolute_import

from behave import given, when, then

from hamcrest import assert_that, equal_to, is_

from pptx import Presentation
from pptx.constants import PP
from pptx.util import Inches

from .helpers import saved_pptx_path


# given ===================================================

@given('I have a reference to a paragraph')
def step_given_ref_to_paragraph(context):
    context.prs = Presentation()
    blank_slidelayout = context.prs.slidelayouts[6]
    slide = context.prs.slides.add_slide(blank_slidelayout)
    length = Inches(2.00)
    textbox = slide.shapes.add_textbox(length, length, length, length)
    context.p = textbox.textframe.paragraphs[0]


@given('I have a reference to a textframe')
def step_given_ref_to_textframe(context):
    context.prs = Presentation()
    blank_slidelayout = context.prs.slidelayouts[6]
    slide = context.prs.slides.add_slide(blank_slidelayout)
    length = Inches(2.00)
    textbox = slide.shapes.add_textbox(length, length, length, length)
    context.textframe = textbox.textframe


# when ====================================================

@when('I indent the first paragraph')
def step_when_indent_first_paragraph(context):
    p = context.body.textframe.paragraphs[0]
    p.level = 1


@when("I set the paragraph alignment to centered")
def step_when_set_paragraph_alignment_to_centered(context):
    context.p.alignment = PP.ALIGN_CENTER


@when("I set the textframe word wrap to True")
def step_when_set_textframe_word_wrap_to_true(context):
    context.textframe.word_wrap = True


@when("I set the textframe word wrap to False")
def step_when_set_textframe_word_wrap_to_false(context):
    context.textframe.word_wrap = False


@when("I set the textframe word wrap to None")
def step_when_set_textframe_word_wrap_to_none(context):
    context.textframe.word_wrap = None


# then ====================================================

@then('the paragraph is aligned centered')
def step_then_paragraph_is_aligned_centered(context):
    prs = Presentation(saved_pptx_path)
    p = prs.slides[0].shapes[0].textframe.paragraphs[0]
    assert_that(p.alignment, is_(equal_to(PP.ALIGN_CENTER)))


@then('the paragraph is indented to the second level')
def step_then_paragraph_indented_to_second_level(context):
    prs = Presentation(saved_pptx_path)
    sld = prs.slides[0]
    body = sld.shapes.placeholders[1]
    p = body.textframe.paragraphs[0]
    assert_that(p.level, is_(equal_to(1)))


@then('the textframe word wrap is empty')
def step_them_textframe_word_wrap_is_empty(context):
    prs = Presentation(saved_pptx_path)
    textframe = prs.slides[0].shapes[0].textframe
    assert_that(textframe.word_wrap, is_(None))


@then('the textframe word wrap is off')
def step_them_textframe_word_wrap_is_off(context):
    prs = Presentation(saved_pptx_path)
    textframe = prs.slides[0].shapes[0].textframe
    assert_that(textframe.word_wrap, is_(False))


@then('the textframe word wrap is on')
def step_them_textframe_word_wrap_is_on(context):
    prs = Presentation(saved_pptx_path)
    textframe = prs.slides[0].shapes[0].textframe
    assert_that(textframe.word_wrap, is_(True))
