# encoding: utf-8

"""
Gherkin step implementations for chart features.
"""

from __future__ import absolute_import, print_function

from behave import given, then, when

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


@given('a bar plot {having_or_not} data labels')
def given_a_bar_plot_having_or_not_data_labels(context, having_or_not):
    slide_idx = {
        'having':     0,
        'not having': 1,
    }[having_or_not]
    prs = Presentation(test_pptx('cht-plot-props'))
    context.plot = prs.slides[slide_idx].shapes[0].chart.plots[0]


# when ====================================================

@when('I assign {value} to plot.has_data_labels')
def when_I_assign_value_to_plot_has_data_labels(context, value):
    new_value = {
        'True':  True,
        'False': False,
    }[value]
    context.plot.has_data_labels = new_value


# then ====================================================

@then('I can access the chart category axis')
def then_I_can_access_the_chart_category_axis(context):
    category_axis = context.chart.category_axis
    assert isinstance(category_axis, CategoryAxis)


@then('I can access the chart value axis')
def then_I_can_access_the_chart_value_axis(context):
    value_axis = context.chart.value_axis
    assert isinstance(value_axis, ValueAxis)


@then('the plot.has_data_labels property is {value}')
def then_the_plot_has_data_labels_property_is_value(context, value):
    expected_value = {
        'True':  True,
        'False': False,
    }[value]
    assert context.plot.has_data_labels is expected_value
