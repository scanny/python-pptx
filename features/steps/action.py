# encoding: utf-8

"""
Gherkin step implementations for click action-related features.
"""

from __future__ import absolute_import, print_function

from behave import given, then

from pptx import Presentation
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
    sld = Presentation(test_pptx('act-props')).slides[0]
    context.shape = sld.shapes[shape_idx]


# when ====================================================


# then ====================================================

@then('click_action.action is {member_name}')
def then_click_action_action_is_value(context, member_name):
    click_action = context.shape.click_action
    expected_value = getattr(PP_ACTION, member_name)
    assert click_action.action == expected_value
