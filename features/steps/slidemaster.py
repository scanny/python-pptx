# encoding: utf-8

"""
Step implementations for slide master-related features
"""

from __future__ import absolute_import

from behave import given, then

from pptx import Presentation
from pptx.parts.slidemaster import _SlideLayouts
from pptx.parts.slides import SlideLayout

from .helpers import test_pptx


# given ===================================================

@given('a slide master having two slide layouts')
def given_slide_master_having_two_layouts(context):
    prs = Presentation(test_pptx('mst-slide-layouts'))
    context.slide_master = prs.slide_master


@given('a slide layout collection containing two layouts')
def given_slide_layout_collection_containing_two_layouts(context):
    prs = Presentation(test_pptx('mst-slide-layouts'))
    context.slide_layouts = prs.slide_master.slide_layouts


# then ====================================================

@then('I can access a slide layout by index')
def then_can_access_slide_layout_by_index(context):
    slide_layouts = context.slide_layouts
    for idx in range(2):
        slide_layout = slide_layouts[idx]
        assert isinstance(slide_layout, SlideLayout)


@then('I can access the slide layouts of the slide master')
def then_can_access_slide_layouts_of_slide_master(context):
    slide_master = context.slide_master
    slide_layouts = slide_master.slide_layouts
    msg = 'SlideMaster.slide_layouts not instance of _SlideLayouts'
    assert isinstance(slide_layouts, _SlideLayouts), msg


@then('I can iterate over the slide layouts')
def then_can_iterate_over_the_slide_layouts(context):
    slide_layouts = context.slide_layouts
    actual_count = 0
    for slide_layout in slide_layouts:
        actual_count += 1
        assert isinstance(slide_layout, SlideLayout)
    assert actual_count == 2


@then('the length of the slide layout collection is 2')
def then_len_of_slide_layout_collection_is_2(context):
    slide_master = context.slide_master
    slide_layouts = slide_master.slide_layouts
    assert len(slide_layouts) == 2, (
        'expected len(slide_layouts) of 2, got %s' % len(slide_layouts)
    )
