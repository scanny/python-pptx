# encoding: utf-8

"""
Gherkin step implementations for click action-related features.
"""

from __future__ import absolute_import, print_function

from behave import given, then, when

from pptx import Presentation
from pptx.action import Hyperlink
from pptx.enum.action import PP_ACTION

from helpers import test_pptx


# given ===================================================

@given('a shape having click action {action}')
def given_a_shape_having_click_action_action(context, action):
    shape_idx = (
        'none',
        'first slide',
        'last slide',
        'previous slide',
        'next slide',
        'last slide viewed',
        'named slide',
        'end show',
        'hyperlink',
        'other presentation',
        'open file',
        'custom slide show',
        'OLE action',
        'run macro',
        'run program',
    ).index(action)
    slides = Presentation(test_pptx('act-props')).slides
    context.slides = slides
    context.shape = slides[2].shapes[shape_idx]


# when ====================================================

@when('I assign {value} to click_action.hyperlink.address')
def when_I_assign_value_to_click_action_hyperlink_address(context, value):
    value = None if value == 'None' else value
    context.shape.click_action.hyperlink.address = value


# then ====================================================

@then('click_action.action is {member_name}')
def then_click_action_action_is_value(context, member_name):
    click_action = context.shape.click_action
    expected_value = getattr(PP_ACTION, member_name)
    assert click_action.action == expected_value


@then('click_action.hyperlink is a Hyperlink object')
def then_click_action_hyperlink_is_a_Hyperlink_object(context):
    hyperlink = context.shape.click_action.hyperlink
    assert isinstance(hyperlink, Hyperlink)


@then('click_action.hyperlink.address is {value}')
def then_click_action_hyperlink_address_is_value(context, value):
    expected_value = None if value == 'None' else value
    hyperlink = context.shape.click_action.hyperlink
    print('expected value %s != %s' % (expected_value, hyperlink.address))
    assert hyperlink.address == expected_value


@then('click_action.target_slide is slide {idx}')
def then_click_action_target_slide_is_slide_idx(context, idx):
    expected_value = None if idx == 'None' else context.slides[int(idx)]
    click_action = context.shape.click_action
    assert click_action.target_slide == expected_value
