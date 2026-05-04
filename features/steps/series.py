"""Gherkin step implementations for chart plot features."""

from __future__ import annotations

from ast import literal_eval

from behave import given, then, when
from helpers import test_pptx

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.chart import (
    XL_ERROR_BAR_INCLUDE,
    XL_ERROR_BAR_TYPE,
    XL_MARKER_STYLE,
    XL_TRENDLINE_TYPE,
)
from pptx.enum.dml import MSO_FILL_TYPE, MSO_THEME_COLOR

# given ===================================================


@given("a BarSeries object having {fill_type} fill as series")
def given_a_BarSeries_object_having_fill_type_as_series(context, fill_type):
    series_idx = {"Automatic": 0, "No Fill": 1, "Orange": 2, "Accent 1": 3}[fill_type]
    prs = Presentation(test_pptx("cht-series"))
    plot = prs.slides[2].shapes[0].chart.plots[0]
    context.series = plot.series[series_idx]


@given("a BarSeries object having invert_if_negative of {setting} as series")
def given_a_bar_series_having_invert_if_negative_setting(context, setting):
    series_idx = {"no explicit setting": 0, "True": 1, "False": 2}[setting]
    prs = Presentation(test_pptx("cht-series"))
    plot = prs.slides[2].shapes[0].chart.plots[0]
    context.series = plot.series[series_idx]


@given("a BarSeries object having values {values} as series")
def given_a_bar_series_having_values_as_series(context, values):
    prs = Presentation(test_pptx("cht-series"))
    series_idx = {"1.2, 2.3, 3.4": 0, "4.5, None, 6.7": 1}[values]
    context.series = prs.slides[3].shapes[0].chart.plots[0].series[series_idx]


@given("a BarSeries object having {width} line as series")
def given_a_bar_series_having_width_line_as_series(context, width):
    series_idx = {"no": 0, "1 point": 1}[width]
    prs = Presentation(test_pptx("cht-series"))
    plot = prs.slides[2].shapes[0].chart.plots[0]
    context.series = plot.series[series_idx]


@given("a marker")
def given_a_marker(context):
    prs = Presentation(test_pptx("cht-marker-props"))
    series = prs.slides[0].shapes[0].chart.series[0]
    context.marker = series.marker


@given("a marker having size of {case}")
def given_a_marker_having_size_of_case(context, case):
    series_idx = {"no explicit value": 0, "24 points": 1, "36 points": 2}[case]
    prs = Presentation(test_pptx("cht-marker-props"))
    series = prs.slides[0].shapes[0].chart.series[series_idx]
    context.marker = series.marker


@given("a marker having style of {case}")
def given_a_marker_having_style_of_case(context, case):
    series_idx = {"no explicit value": 0, "circle": 1, "triangle": 2}[case]
    prs = Presentation(test_pptx("cht-marker-props"))
    series = prs.slides[0].shapes[0].chart.series[series_idx]
    context.marker = series.marker


@given("a point")
def given_a_point(context):
    prs = Presentation(test_pptx("cht-point-props"))
    chart = prs.slides[0].shapes[0].chart
    context.point = chart.plots[0].series[0].points[0]


@given("a {points_type} object containing 3 points")
def given_a_points_type_object_containing_3_points(context, points_type):
    slide_idx = {"XyPoints": 0, "BubblePoints": 1, "CategoryPoints": 2}[points_type]
    prs = Presentation(test_pptx("cht-point-access"))
    series = prs.slides[slide_idx].shapes[0].chart.plots[0].series[0]
    context.points = series.points


@given("a series")
def given_a_series(context):
    prs = Presentation(test_pptx("cht-series"))
    context.series = prs.slides[3].shapes[0].chart.plots[0].series[0]


@given("a {prefix}Series object as series")
def given_a_series_of_type_series_type(context, prefix):
    slide_idx = {
        "Area": 8,
        "Bar": 3,
        "Bubble": 5,
        "Category": 3,
        "Doughnut": 9,
        "Line": 6,
        "Pie": 10,
        "Radar": 7,
        "Xy": 4,
    }[prefix]
    prs = Presentation(test_pptx("cht-series"))
    context.series = prs.slides[slide_idx].shapes[0].chart.plots[0].series[0]


@given("a SeriesCollection object for a plot having {n} series")
def given_a_SeriesCollection_object_for_a_plot_having_n_series(context, n):
    prs = Presentation(test_pptx("cht-series"))
    plot = prs.slides[0].shapes[0].chart.plots[0]
    context.series_collection = plot.series
    context.series_count = int(n)


@given("a SeriesCollection object for a {type_} chart having {n} series")
def given_a_SeriesCollection_for_chart_having_n_series(context, type_, n):
    slide_idx = {"single-plot": 0, "multi-plot": 1}[type_]
    prs = Presentation(test_pptx("cht-series"))
    context.series_collection = prs.slides[slide_idx].shapes[0].chart.series
    context.series_count = int(n)


# when ====================================================


@when("I add a series with number format {strval}")
def when_I_add_a_series_with_number_format(context, strval):
    chart_data = context.chart_data
    params = {"name": "Series Foo"}
    if strval != "None":
        params["number_format"] = int(strval)
    context.series_data = chart_data.add_series(**params)


@when("I assign {value} to marker.size")
def when_I_assign_value_to_marker_size(context, value):
    new_value = None if value == "None" else int(value)
    context.marker.size = new_value


@when("I assign {value} to marker.style")
def when_I_assign_value_to_marker_style(context, value):
    new_value = None if value == "None" else getattr(XL_MARKER_STYLE, value)
    context.marker.style = new_value


@when("I assign {value} to point.invert_if_negative")
def when_I_assign_value_to_point_invert_if_negative(context, value):
    new_value = {"True": True, "False": False}[value]
    context.point.invert_if_negative = new_value


@when("I assign {value} to series.invert_if_negative")
def when_I_assign_value_to_series_invert_if_negative(context, value):
    new_value = {"True": True, "False": False}[value]
    context.series.invert_if_negative = new_value


# then ====================================================


@then("data_point.number_format is {value_str}")
def then_data_point_number_format_is(context, value_str):
    data_point = context.data_point
    number_format = value_str if value_str == "General" else int(value_str)
    assert data_point.number_format == number_format


@then("iterating points produces 3 Point objects")
def then_iterating_points_produces_3_point_objects(context):
    points = context.points
    idx = -1
    for idx, point in enumerate(points):
        assert type(point).__name__ == "Point"
    assert idx == 2, "got %s" % idx


@then("iterating series_collection produces {count} Series objects")
def then_iterating_series_collection_produces_count_series(context, count):
    expected_idx = int(count) - 1
    idx = -1
    for idx, series in enumerate(context.series_collection):
        type_name = type(series).__name__
        assert type_name.endswith("Series"), "got %s" % type_name
    assert idx == expected_idx, "got %s" % idx


@then("len(points) is 3")
def then_len_points_is_3(context):
    points = context.points
    assert len(points) == 3


@then("len(series_collection) is {count}")
def then_len_series_collection_is_count(context, count):
    expected_len = int(count)
    actual_len = len(context.series_collection)
    assert actual_len == expected_len, "got %s" % actual_len


@then("len(series.values) is {count} for each series")
def then_len_series_values_is_count_for_each_series(context, count):
    expected_count = int(count)
    for series in context.chart.plots[0].series:
        assert len(series.values) == expected_count


@then("marker.format is a ChartFormat object")
def then_marker_format_is_a_ChartFormat_object(context):
    marker = context.marker
    assert type(marker.format).__name__ == "ChartFormat"


@then("marker.format.fill is a FillFormat object")
def then_marker_format_fill_is_a_FillFormat_object(context):
    marker = context.marker
    assert type(marker.format.fill).__name__ == "FillFormat"


@then("marker.format.line is a LineFormat object")
def then_marker_format_line_is_a_LineFormat_object(context):
    marker = context.marker
    assert type(marker.format.line).__name__ == "LineFormat"


@then("marker.size is {case}")
def then_marker_size_is_case(context, case):
    expected_value = None if case == "None" else int(case)
    marker = context.marker
    assert marker.size == expected_value, "got %s" % marker.size


@then("marker.style is {case}")
def then_marker_style_is_case(context, case):
    expected_value = None if case == "None" else getattr(XL_MARKER_STYLE, case)
    marker = context.marker
    assert marker.style == expected_value, "got %s" % marker.style


@then("point.data_label is a DataLabel object")
def then_point_data_label_is_a_DataLabel_object(context):
    point = context.point
    assert type(point.data_label).__name__ == "DataLabel"


@then("point.invert_if_negative is {value}")
def then_point_invert_if_negative_is_value(context, value):
    expected_value = {"True": True, "False": False}[value]
    actual_value = context.point.invert_if_negative
    assert actual_value is expected_value, "got %s" % actual_value


@then("point.format is a ChartFormat object")
def then_point_format_is_a_ChartFormat_object(context):
    point = context.point
    assert type(point.format).__name__ == "ChartFormat"


@then("point.format.fill is a FillFormat object")
def then_point_format_fill_is_a_FillFormat_object(context):
    point = context.point
    assert type(point.format.fill).__name__ == "FillFormat"


@then("point.format.line is a LineFormat object")
def then_point_format_line_is_a_LineFormat_object(context):
    point = context.point
    assert type(point.format.line).__name__ == "LineFormat"


@then("point.marker is a Marker object")
def then_point_marker_is_a_Marker_object(context):
    point = context.point
    assert type(point.marker).__name__ == "Marker"


@then("points[2] is a Point object")
def then_points_2_is_a_Point_object(context):
    actual = type(context.points[2]).__name__
    assert actual == "Point", "points[2] is a %s object" % actual


@then("series_collection[2] is a Series object")
def then_series_collection_2_is_a_Series_object(context):
    type_name = type(context.series_collection[2]).__name__
    assert type_name.endswith("Series"), "got %s" % type_name


@then("series.data_labels is a DataLabels object")
def then_series_data_labels_is_a_DataLabels_object(context):
    actual = type(context.series.data_labels).__name__
    assert actual == "DataLabels", "series.data_labels is a %s object" % actual


@then("series.format.fill.fore_color.rgb is FF6600")
def then_series_format_fill_fore_color_rgb_is_FF6600(context):
    rgb_color = context.series.format.fill.fore_color.rgb
    assert rgb_color == RGBColor(0xFF, 0x66, 0x00), "got %s" % rgb_color


@then("series.format.fill.fore_color.theme_color is Accent 1")
def then_series_format_fill_fore_color_theme_color_is_Accent_1(context):
    theme_color = context.series.format.fill.fore_color.theme_color
    assert theme_color == MSO_THEME_COLOR.ACCENT_1, "got %s" % theme_color


@then("series.format.fill.type is {fill_type}")
def then_series_format_fill_type_is_type(context, fill_type):
    expected_fill_type = {
        "None": None,
        "MSO_FILL_TYPE.BACKGROUND": MSO_FILL_TYPE.BACKGROUND,
        "MSO_FILL_TYPE.SOLID": MSO_FILL_TYPE.SOLID,
    }[fill_type]
    fill_type = context.series.format.fill.type
    assert fill_type == expected_fill_type, "got %s" % fill_type


@then("series.invert_if_negative is {value}")
def then_series_invert_if_negative_is_value(context, value):
    expected_value = {"True": True, "False": False}[value]
    series = context.series
    assert series.invert_if_negative is expected_value


@then("series.format.line.width is {width}")
def then_series_format_line_width_is_width(context, width):
    expected_width = int(width)
    line_width = context.series.format.line.width
    assert line_width == expected_width, "got %s" % line_width


@then("series.format is a ChartFormat object")
def then_series_format_is_a_ChartFormat_object(context):
    actual = type(context.series.format).__name__
    assert actual == "ChartFormat", "series.format is a %s object" % actual


@then("series.marker is a Marker object")
def then_series_marker_is_a_Marker_object(context):
    actual = type(context.series.marker).__name__
    assert actual == "Marker", "series.marker is a %s object" % actual


@then("series.points is a {type_name} object")
def then_series_points_is_a_type_name_object(context, type_name):
    actual = type(context.series.points).__name__
    expected = type_name
    assert actual == expected, "series.points is a %s object" % actual


@then("series.values is {values}")
def then_series_values_is_values(context, values):
    series = context.series
    expected_values = literal_eval(values)
    assert series.values == expected_values, "got %s" % (series.values,)


# error-bar steps -----------------------------------------


@given("series has fixed-value error bars attached")
def given_series_has_fixed_value_error_bars_attached(context):
    context.series.set_error_bars(
        type_=XL_ERROR_BAR_TYPE.FIXED_VALUE,
        value=1.5,
        include=XL_ERROR_BAR_INCLUDE.BOTH,
    )


@when("I call series.set_error_bars with type {type_name} and value {value}")
def when_I_call_series_set_error_bars(context, type_name, value):
    context.series.set_error_bars(
        type_=getattr(XL_ERROR_BAR_TYPE, type_name),
        value=float(value),
    )


@when("I assign None to series.error_bars")
def when_I_assign_None_to_series_error_bars(context):
    context.series.error_bars = None


@then("series.has_error_bars is {value}")
def then_series_has_error_bars_is_value(context, value):
    expected = {"True": True, "False": False}[value]
    actual = context.series.has_error_bars
    assert actual is expected, "got %s" % actual


@then("series.error_bars is None")
def then_series_error_bars_is_None(context):
    assert context.series.error_bars is None, "got %s" % context.series.error_bars


@then("series.error_bars.type is {type_name}")
def then_series_error_bars_type_is(context, type_name):
    expected = getattr(XL_ERROR_BAR_TYPE, type_name)
    actual = context.series.error_bars.type
    assert actual == expected, "got %s" % actual


@then("series.error_bars.value is {value}")
def then_series_error_bars_value_is(context, value):
    expected = float(value)
    actual = context.series.error_bars.value
    assert actual == expected, "got %s" % actual


@then("series.error_bars.include is {include_name}")
def then_series_error_bars_include_is(context, include_name):
    expected = getattr(XL_ERROR_BAR_INCLUDE, include_name)
    actual = context.series.error_bars.include
    assert actual == expected, "got %s" % actual


# trendline steps -----------------------------------------


@given("series has a linear trendline attached")
def given_series_has_a_linear_trendline_attached(context):
    context.series.add_trendline(XL_TRENDLINE_TYPE.LINEAR)


@when("I call series.add_trendline with type {type_name} and display equation on")
def when_I_call_series_add_trendline_with_display_equation(context, type_name):
    context.series.add_trendline(
        getattr(XL_TRENDLINE_TYPE, type_name),
        display_equation=True,
    )


@when("I call series.add_trendline with type {type_name} and order {order}")
def when_I_call_series_add_trendline_with_order(context, type_name, order):
    context.series.add_trendline(
        getattr(XL_TRENDLINE_TYPE, type_name),
        order=int(order),
    )


@when("I call series.add_trendline with type {type_name} and period {period}")
def when_I_call_series_add_trendline_with_period(context, type_name, period):
    context.series.add_trendline(
        getattr(XL_TRENDLINE_TYPE, type_name),
        period=int(period),
    )


@when("I call delete() on series.trendlines[0]")
def when_I_call_delete_on_series_trendlines_0(context):
    context.series.trendlines[0].delete()


@then("series.trendlines is an empty list")
def then_series_trendlines_is_an_empty_list(context):
    assert context.series.trendlines == [], "got %s" % context.series.trendlines


@then("series.trendlines has length {count:d}")
def then_series_trendlines_has_length(context, count):
    actual = len(context.series.trendlines)
    assert actual == count, "got %d" % actual


@then("series.trendlines[0].trendline_type is {type_name}")
def then_series_trendlines_0_trendline_type_is(context, type_name):
    expected = getattr(XL_TRENDLINE_TYPE, type_name)
    actual = context.series.trendlines[0].trendline_type
    assert actual == expected, "got %s" % actual


@then("series.trendlines[0].display_equation is {value}")
def then_series_trendlines_0_display_equation_is(context, value):
    expected = {"True": True, "False": False}[value]
    actual = context.series.trendlines[0].display_equation
    assert actual is expected, "got %s" % actual


@then("series.trendlines[0].display_r_squared is {value}")
def then_series_trendlines_0_display_r_squared_is(context, value):
    expected = {"True": True, "False": False}[value]
    actual = context.series.trendlines[0].display_r_squared
    assert actual is expected, "got %s" % actual


@then("series.trendlines[0].order is {value:d}")
def then_series_trendlines_0_order_is(context, value):
    actual = context.series.trendlines[0].order
    assert actual == value, "got %s" % actual


@then("series.trendlines[0].period is {value:d}")
def then_series_trendlines_0_period_is(context, value):
    actual = context.series.trendlines[0].period
    assert actual == value, "got %s" % actual


# source-range / category-range / name-range steps --------


@when("I assign a new source_range formula to series.source_range")
def when_I_assign_new_source_range(context):
    context.new_source_range = "Sheet1!$B$2:$Z$2"
    context.series.source_range = context.new_source_range


@then("series.source_range is a non-empty string")
def then_series_source_range_is_a_non_empty_string(context):
    actual = context.series.source_range
    assert isinstance(actual, str), "got %r" % (actual,)
    assert actual, "source_range is empty"


@then("series.source_range ends with the value-range A1 reference")
def then_series_source_range_ends_with_a1(context):
    # -- whatever sheet name the fixture uses, the A1-portion must be present --
    ref = context.series.values_sheet_reference
    assert ref is not None, "values_sheet_reference is None"
    assert "!" in context.series.source_range
    assert ref.a1_range, "a1_range is empty"


@then("series.values_sheet_reference.sheet_name is a non-empty string")
def then_series_values_sheet_reference_sheet_name(context):
    ref = context.series.values_sheet_reference
    assert ref is not None, "values_sheet_reference is None"
    assert isinstance(ref.sheet_name, str), "got %r" % (ref.sheet_name,)
    assert ref.sheet_name, "sheet_name is empty"


@then("series.category_range is a non-empty string")
def then_series_category_range_is_a_non_empty_string(context):
    actual = context.series.category_range
    assert isinstance(actual, str), "got %r" % (actual,)
    assert actual, "category_range is empty"


@then("series.name_range is a non-empty string")
def then_series_name_range_is_a_non_empty_string(context):
    actual = context.series.name_range
    assert isinstance(actual, str), "got %r" % (actual,)
    assert actual, "name_range is empty"


@then("series.source_range reflects the new formula")
def then_series_source_range_reflects_new(context):
    assert context.series.source_range == context.new_source_range, (
        "got %r" % context.series.source_range
    )


# series.delete() (issue #1043) ----------------------------


@given("a chart with 5 series named S0, S1, S2, S3, S4")
def given_a_chart_with_5_series_S0_S4(context):
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    cd = CategoryChartData()
    cd.categories = ["A", "B", "C"]
    for i in range(5):
        cd.add_series("S%d" % i, (float(i), float(i + 1), float(i + 2)))
    gf = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        cd,
    )
    context.chart = gf.chart


@when("I delete chart.series[3] and chart.series[1]")
def when_I_delete_chart_series_3_and_1(context):
    context.chart.series[3].delete()
    context.chart.series[1].delete()


@then("chart.series has length {count:d}")
def then_chart_series_has_length(context, count):
    actual = len(context.chart.series)
    assert actual == count, "got %d" % actual


@then("the remaining series names are S0, S2, S4")
def then_the_remaining_series_names_are_S0_S2_S4(context):
    names = [s.name for s in context.chart.series]
    assert names == ["S0", "S2", "S4"], "got %r" % names
