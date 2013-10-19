# encoding: utf-8

"""
Gherkin step implementations for placeholder-related features.
"""

import os

from behave import given, when, then
from hamcrest import assert_that, equal_to, is_

from pptx import Presentation


def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))

thisdir = os.path.split(__file__)[0]
scratch_dir = absjoin(thisdir, '../_scratch')
saved_pptx_path = absjoin(scratch_dir, 'test_out.pptx')

test_text = "python-pptx was here!"


# given ===================================================

@given('I have a reference to a bullet body placeholder')
def step_given_ref_to_bullet_body_placeholder(context):
    context.prs = Presentation()
    slidelayout = context.prs.slidelayouts[1]
    context.sld = context.prs.slides.add_slide(slidelayout)
    context.body = context.sld.shapes.placeholders[1]


# when ====================================================

@when("I set the title text of the slide")
def step_when_set_slide_title_text(context):
    context.sld.shapes.title.text = test_text


# then ====================================================

@then('the text appears in the title placeholder')
def step_then_text_appears_in_title_placeholder(context):
    prs = Presentation(saved_pptx_path)
    title_shape = prs.slides[0].shapes.title
    title_text = title_shape.textframe.paragraphs[0].runs[0].text
    assert_that(title_text, is_(equal_to(test_text)))
