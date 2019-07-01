# encoding: utf-8

"""Gherkin step implementations for slide background-related features."""

from __future__ import absolute_import, division, print_function, unicode_literals

from behave import given, then

from pptx import Presentation

from helpers import test_pptx


# given ===================================================


@given("a _Background object having {type} background as background")
def given_a_Background_object_having_type_background(context, type):
    sld_idx = {"no": 0, "a fill": 1, "a style reference": None}[type]
    prs = Presentation(test_pptx("sld-background"))
    slide = prs.slide_masters[0] if sld_idx is None else prs.slides[sld_idx]
    context.background = slide.background


# then ====================================================


@then("background.fill is a FillFormat object")
def then_background_fill_is_a_Fill_object(context):
    cls_name = context.background.fill.__class__.__name__
    assert cls_name == "FillFormat", "background.fill is a %s object" % cls_name
