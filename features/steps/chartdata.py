# encoding: utf-8

"""Gherkin step implementations for chart data features."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import datetime

from behave import given, then, when

from pptx.chart.data import (
    BubbleChartData, Category, CategoryChartData, XyChartData
)
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches


# given ===================================================

@given('a BubbleChartData object with number format {strval}')
def given_a_BubbleChartData_object_with_number_format(context, strval):
    params = {}
    if strval != 'None':
        params['number_format'] = int(strval)
    context.chart_data = BubbleChartData(**params)


@given('a Categories object with number format {init_nf}')
def given_a_Categories_object_with_number_format_init_nf(context, init_nf):
    categories = CategoryChartData().categories
    if init_nf != 'left as default':
        categories.number_format = init_nf
    context.categories = categories


@given('a Category object')
def given_a_Category_object(context):
    context.category = Category(None, None)


@given('a CategoryChartData object')
def given_a_CategoryChartData_object(context):
    context.chart_data = CategoryChartData()


@given('a CategoryChartData object having date categories')
def given_a_CategoryChartData_object_having_date_categories(context):
    chart_data = CategoryChartData()
    chart_data.categories = [
        datetime.date(2016, 12, 27),
        datetime.date(2016, 12, 28),
        datetime.date(2016, 12, 29),
    ]
    context.chart_data = chart_data


@given('a CategoryChartData object with number format {strval}')
def given_a_CategoryChartData_object_with_number_format(context, strval):
    params = {}
    if strval != 'None':
        params['number_format'] = int(strval)
    context.chart_data = CategoryChartData(**params)


@given('a XyChartData object with number format {strval}')
def given_a_XyChartData_object_with_number_format(context, strval):
    params = {}
    if strval != 'None':
        params['number_format'] = int(strval)
    context.chart_data = XyChartData(**params)


@given('the categories are of type {type_}')
def given_the_categories_are_of_type(context, type_):
    label = {
        'date':  datetime.date(2016, 12, 22),
        'float': 42.24,
        'int':   42,
        'str':   'foobar',
    }[type_]
    context.categories.add_category(label)


# when ====================================================

@when('I add a bubble data point with number format {strval}')
def when_I_add_a_bubble_data_point_with_number_format(context, strval):
    series_data = context.series_data
    params = {'x': 1, 'y': 2, 'size': 10}
    if strval != 'None':
        params['number_format'] = int(strval)
    context.data_point = series_data.add_data_point(**params)


@when('I add a data point with number format {strval}')
def when_I_add_a_data_point_with_number_format(context, strval):
    series_data = context.series_data
    params = {'value': 42}
    if strval != 'None':
        params['number_format'] = int(strval)
    context.data_point = series_data.add_data_point(**params)


@when('I add an XY data point with number format {strval}')
def when_I_add_an_XY_data_point_with_number_format(context, strval):
    series_data = context.series_data
    params = {'x': 1, 'y': 2}
    if strval != 'None':
        params['number_format'] = int(strval)
    context.data_point = series_data.add_data_point(**params)


@when('I add an {xy_type} chart having 2 series of 3 points each')
def when_I_add_an_xy_chart_having_2_series_of_3_points(context, xy_type):
    chart_type = getattr(XL_CHART_TYPE, xy_type)
    data = (
        ('Series 1', ((-0.1, 0.5), (16.2, 0.0), (8.0,  0.2))),
        ('Series 2', ((12.4, 0.8), (-7.5, -0.5), (-5.1, -0.2)))
    )

    chart_data = XyChartData()

    for series_data in data:
        series_label, points = series_data
        series = chart_data.add_series(series_label)
        for point in points:
            x, y = point
            series.add_data_point(x, y)

    context.chart = context.slide.shapes.add_chart(
        chart_type, Inches(1), Inches(1), Inches(8), Inches(5), chart_data
    ).chart


@when("I assign ['a', 'b', 'c'] to chart_data.categories")
def when_I_assign_a_b_c_to_chart_data_categories(context):
    chart_data = context.chart_data
    chart_data.categories = ['a', 'b', 'c']


# then ====================================================

@then("[c.label for c in chart_data.categories] is ['a', 'b', 'c']")
def then_c_label_for_c_in_chart_data_categories_is_a_b_c(context):
    chart_data = context.chart_data
    assert [c.label for c in chart_data.categories] == ['a', 'b', 'c']


@then('categories.number_format is {value}')
def then_categories_number_format_is_value(context, value):
    expected_value = value
    number_format = context.categories.number_format
    assert number_format == expected_value, 'got %s' % number_format


@then('category.add_sub_category(name) is a Category object')
def then_category_add_sub_category_is_a_Category_object(context):
    category = context.category
    context.sub_category = sub_category = category.add_sub_category('foobar')
    assert type(sub_category).__name__ == 'Category'


@then('category.sub_categories[-1] is the new category')
def then_category_sub_categories_minus_1_is_the_new_category(context):
    category, sub_category = context.category, context.sub_category
    assert category.sub_categories[-1] is sub_category


@then('chart_data.add_category(name) is a Category object')
def then_chart_data_add_category_name_is_a_Category_object(context):
    chart_data = context.chart_data
    context.category = category = chart_data.add_category('foobar')
    assert type(category).__name__ == 'Category'


@then('chart_data.add_series(name, values) is a CategorySeriesData object')
def then_chart_data_add_series_is_a_CategorySeriesData_object(context):
    chart_data = context.chart_data
    context.series = series = chart_data.add_series('Series X', (1, 2, 3))
    assert type(series).__name__ == 'CategorySeriesData'


@then('chart_data.categories is a Categories object')
def then_chart_data_categories_is_a_Categories_object(context):
    chart_data = context.chart_data
    assert type(chart_data.categories).__name__ == 'Categories'


@then('chart_data.categories[-1] is the category')
def then_chart_data_categories_minus_1_is_the_category(context):
    chart_data, category = context.chart_data, context.category
    assert chart_data.categories[-1] is category


@then('chart_data.number_format is {value_str}')
def then_chart_data_number_format_is(context, value_str):
    chart_data = context.chart_data
    number_format = value_str if value_str == 'General' else int(value_str)
    assert chart_data.number_format == number_format


@then('chart_data[-1] is the new series')
def then_chart_data_minus_1_is_the_new_series(context):
    chart_data, series = context.chart_data, context.series
    assert chart_data[-1] is series


@then('series_data.number_format is {value_str}')
def then_series_data_number_format_is(context, value_str):
    series_data = context.series_data
    number_format = value_str if value_str == 'General' else int(value_str)
    assert series_data.number_format == number_format
