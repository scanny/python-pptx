# encoding: utf-8

"""Gherkin step implementations for chart data label features."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from behave import given, then, when

from pptx import Presentation
from pptx.enum.chart import XL_DATA_LABEL_POSITION

from helpers import test_pptx


# given ===================================================

@given('a data label {having_or_not} custom font')
def given_a_data_label_having_or_not_custom_font(context, having_or_not):
    point_idx = {
        'having a':  0,
        'having no': 1,
    }[having_or_not]
    prs = Presentation(test_pptx('cht-point-props'))
    points = prs.slides[2].shapes[0].chart.plots[0].series[0].points
    context.data_label = points[point_idx].data_label


@given('a data label {having_or_not} custom text')
def given_a_data_label_having_or_not_custom_text(context, having_or_not):
    point_idx = {
        'having':    0,
        'having no': 1,
    }[having_or_not]
    prs = Presentation(test_pptx('cht-point-props'))
    plot = prs.slides[0].shapes[0].chart.plots[0]
    context.data_label = plot.series[0].points[point_idx].data_label


@given('a data label positioned {relation} its data point')
def given_a_data_label_positioned_relation_its_data_point(context, relation):
    point_idx = {
        'in unspecified relation to': 0,
        'centered on':                1,
        'below':                      2,
    }[relation]
    prs = Presentation(test_pptx('cht-point-props'))
    plot = prs.slides[1].shapes[0].chart.plots[0]
    context.data_label = plot.series[0].points[point_idx].data_label


@given('bar chart data labels positioned {relation_to} their data point')
def given_bar_chart_data_labels_positioned_relation_to_their_data_point(
        context, relation_to):
    slide_idx = {
        'in unspecified relation to': 0,
        'inside, at the base of':     1,
    }[relation_to]
    prs = Presentation(test_pptx('cht-datalabels-props'))
    chart = prs.slides[slide_idx].shapes[0].chart
    context.data_labels = chart.plots[0].data_labels


# when ====================================================

@when('I assign {value} to data_label.has_text_frame')
def when_I_assign_value_to_data_label_has_text_frame(context, value):
    new_value = {'True': True, 'False': False}[value]
    context.data_label.has_text_frame = new_value


@when('I assign {value} to data_label.position')
def when_I_assign_value_to_data_label_position(context, value):
    new_value = (
        None if value == 'None' else getattr(XL_DATA_LABEL_POSITION, value)
    )
    context.data_label.position = new_value


@when('I assign {value} to data_labels.position')
def when_I_assign_value_to_data_labels_position(context, value):
    new_value = (
        None if value == 'None' else getattr(XL_DATA_LABEL_POSITION, value)
    )
    context.data_labels.position = new_value


# then ====================================================

@then('data_label.font is a Font object')
def then_data_label_font_is_a_Font_object(context):
    font = context.data_label.font
    assert type(font).__name__ == 'Font'


@then('data_label.has_text_frame is {value}')
def then_data_label_has_text_frame_is_value(context, value):
    expected_value = {'True': True, 'False': False}[value]
    data_label = context.data_label
    assert data_label.has_text_frame is expected_value


@then('data_label.position is {value}')
def then_data_label_position_is_value(context, value):
    expected_value = (
        None if value == 'None' else getattr(XL_DATA_LABEL_POSITION, value)
    )
    data_label = context.data_label
    assert data_label.position is expected_value, (
        'got %s' % data_label.position
    )


@then('data_label.text_frame is a TextFrame object')
def then_data_label_text_frame_is_a_TextFrame_object(context):
    text_frame = context.data_label.text_frame
    assert type(text_frame).__name__ == 'TextFrame'


@then('data_labels.position is {value}')
def then_data_labels_position_is_value(context, value):
    expected_value = (
        None if value == 'None' else getattr(XL_DATA_LABEL_POSITION, value)
    )
    data_labels = context.data_labels
    assert data_labels.position is expected_value, (
        'got %s' % data_labels.position
    )
