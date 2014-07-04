# encoding: utf-8

"""
Gherkin step implementations for chart features.
"""

from __future__ import absolute_import, print_function

from behave import given, then

from pptx import Presentation
from pptx.chart.axis import CategoryAxis, ValueAxis

from .helpers import test_pptx


# given ===================================================

@given('a bar chart')
def given_a_bar_chart(context):
    prs = Presentation(test_pptx('cht-charts'))
    sld = prs.slides[0]
    graphic_frame = sld.shapes[0]
    context.chart = graphic_frame.chart


# then ====================================================

@then('I can access the chart category axis')
def then_I_can_access_the_chart_category_axis(context):
    category_axis = context.chart.category_axis
    assert isinstance(category_axis, CategoryAxis)


@then('I can access the chart value axis')
def then_I_can_access_the_chart_value_axis(context):
    value_axis = context.chart.value_axis
    assert isinstance(value_axis, ValueAxis)
