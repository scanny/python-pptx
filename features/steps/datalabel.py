"""Gherkin step implementations for chart data label features."""

from __future__ import annotations

from behave import given, then, when
from helpers import test_pptx

from pptx import Presentation
from pptx.enum.chart import XL_DATA_LABEL_POSITION

# given ===================================================


@given("a DataLabels object {showing_or_not} category-name as data_labels")
def given_a_DataLabels_object_showing_or_not_cat_name(context, showing_or_not):
    series_idx = {"not showing": 0, "showing": 1}[showing_or_not]
    prs = Presentation(test_pptx("cht-datalabels"))
    chart = prs.slides[2].shapes[0].chart
    context.data_labels = chart.plots[0].series[series_idx].data_labels


@given("a DataLabels object {showing_or_not} legend-key as data_labels")
def given_a_DataLabels_object_showing_or_not_leg_key(context, showing_or_not):
    series_idx = {"not showing": 0, "showing": 1}[showing_or_not]
    prs = Presentation(test_pptx("cht-datalabels"))
    chart = prs.slides[2].shapes[0].chart
    context.data_labels = chart.plots[0].series[series_idx].data_labels


@given("a DataLabels object {showing_or_not} percentage as data_labels")
def given_a_DataLabels_object_showing_or_not_percent(context, showing_or_not):
    slide_idx = {"not showing": 4, "showing": 3}[showing_or_not]
    prs = Presentation(test_pptx("cht-datalabels"))
    chart = prs.slides[slide_idx].shapes[0].chart
    context.data_labels = chart.plots[0].series[0].data_labels


@given("a DataLabels object {showing_or_not} series-name as data_labels")
def given_a_DataLabels_object_showing_or_not_ser_name(context, showing_or_not):
    series_idx = {"not showing": 0, "showing": 1}[showing_or_not]
    prs = Presentation(test_pptx("cht-datalabels"))
    chart = prs.slides[2].shapes[0].chart
    context.data_labels = chart.plots[0].series[series_idx].data_labels


@given("a DataLabels object {showing_or_not} value as data_labels")
def given_a_DataLabels_object_showing_or_not_value(context, showing_or_not):
    series_idx = {"not showing": 0, "showing": 1}[showing_or_not]
    prs = Presentation(test_pptx("cht-datalabels"))
    chart = prs.slides[2].shapes[0].chart
    context.data_labels = chart.plots[0].series[series_idx].data_labels


@given("a DataLabels object with {pos} position as data_labels")
def given_a_DataLabels_object_with_pos_position(context, pos):
    slide_idx = {"inherited": 0, "inside-base": 1}[pos]
    prs = Presentation(test_pptx("cht-datalabels"))
    chart = prs.slides[slide_idx].shapes[0].chart
    context.data_labels = chart.plots[0].data_labels


@given("a data label")
def given_a_data_label(context):
    prs = Presentation(test_pptx("cht-point-props"))
    points = prs.slides[0].shapes[0].chart.plots[0].series[0].points
    context.data_label = points[0].data_label


@given("a data label {having_or_not} custom font as data_label")
def given_a_data_label_having_or_not_custom_font(context, having_or_not):
    point_idx = {"having a": 0, "having no": 1}[having_or_not]
    prs = Presentation(test_pptx("cht-point-props"))
    points = prs.slides[2].shapes[0].chart.plots[0].series[0].points
    context.data_label = points[point_idx].data_label


@given("a data label {having_or_not} custom text as data_label")
def given_a_data_label_having_or_not_custom_text(context, having_or_not):
    point_idx = {"having": 0, "having no": 1}[having_or_not]
    prs = Presentation(test_pptx("cht-point-props"))
    plot = prs.slides[0].shapes[0].chart.plots[0]
    context.data_label = plot.series[0].points[point_idx].data_label


@given("a data label with {pos} position as data_label")
def given_a_data_label_with_pos_position_as_data_label(context, pos):
    point_idx = {"inherited": 0, "centered": 1, "below": 2}[pos]
    prs = Presentation(test_pptx("cht-point-props"))
    plot = prs.slides[1].shapes[0].chart.plots[0]
    context.data_label = plot.series[0].points[point_idx].data_label


# when ====================================================


@when("I assign {value} to data_label.has_text_frame")
def when_I_assign_value_to_data_label_has_text_frame(context, value):
    new_value = {"True": True, "False": False}[value]
    context.data_label.has_text_frame = new_value


@when("I assign '{value}' to data_label.number_format")
def when_I_assign_value_to_data_label_number_format(context, value):
    context.data_label.number_format = value


@when("I assign {value} to data_label.number_format_is_linked")
def when_I_assign_value_to_data_label_number_format_is_linked(context, value):
    context.data_label.number_format_is_linked = {"True": True, "False": False}[value]


@when("I assign {value} to data_label.position")
def when_I_assign_value_to_data_label_position(context, value):
    new_value = None if value == "None" else getattr(XL_DATA_LABEL_POSITION, value)
    context.data_label.position = new_value


@when("I assign {value} to data_labels.position")
def when_I_assign_value_to_data_labels_position(context, value):
    new_value = None if value == "None" else getattr(XL_DATA_LABEL_POSITION, value)
    context.data_labels.position = new_value


@when("I assign {value} to data_labels.show_category_name")
def when_I_assign_value_to_data_labels_show_category_name(context, value):
    context.data_labels.show_category_name = eval(value)


@when("I assign {value} to data_labels.show_legend_key")
def when_I_assign_value_to_data_labels_show_legend_key(context, value):
    context.data_labels.show_legend_key = eval(value)


@when("I assign {value} to data_labels.show_percentage")
def when_I_assign_value_to_data_labels_show_percentage(context, value):
    context.data_labels.show_percentage = eval(value)


@when("I assign {value} to data_labels.show_series_name")
def when_I_assign_value_to_data_labels_show_series_name(context, value):
    context.data_labels.show_series_name = eval(value)


@when("I assign {value} to data_labels.show_value")
def when_I_assign_value_to_data_labels_show_value(context, value):
    context.data_labels.show_value = eval(value)


@when("I assign {value} to data_labels.text_frame.word_wrap")
def when_I_assign_value_to_data_labels_text_frame_word_wrap(context, value):
    new_value = {"True": True, "False": False, "None": None}[value]
    context.data_labels.text_frame.word_wrap = new_value


@when("I set the data label manual layout to ({x:g}, {y:g})")
def when_I_set_the_data_label_manual_layout(context, x, y):
    context.data_label.set_manual_layout(x, y)


@when("I clear the data label manual layout")
def when_I_clear_the_data_label_manual_layout(context):
    context.data_label.clear_manual_layout()


# then ====================================================


@then("data_label.font is a Font object")
def then_data_label_font_is_a_Font_object(context):
    font = context.data_label.font
    assert type(font).__name__ == "Font"


@then("data_label.format is a ChartFormat object")
def then_data_label_format_is_a_ChartFormat_object(context):
    data_label = context.data_label
    assert type(data_label.format).__name__ == "ChartFormat"


@then("data_label.format.fill is a FillFormat object")
def then_data_label_format_fill_is_a_FillFormat_object(context):
    data_label = context.data_label
    assert type(data_label.format.fill).__name__ == "FillFormat"


@then("data_label.format.line is a LineFormat object")
def then_data_label_format_line_is_a_LineFormat_object(context):
    data_label = context.data_label
    assert type(data_label.format.line).__name__ == "LineFormat"


@then("data_label.has_text_frame is {value}")
def then_data_label_has_text_frame_is_value(context, value):
    expected_value = {"True": True, "False": False}[value]
    data_label = context.data_label
    assert data_label.has_text_frame is expected_value


@then("data_label.number_format is '{value}'")
def then_data_label_number_format_is_value(context, value):
    actual = context.data_label.number_format
    assert actual == value, "data_label.number_format is %r, expected %r" % (actual, value)


@then("data_label.number_format_is_linked is {value}")
def then_data_label_number_format_is_linked_is_value(context, value):
    expected = {"True": True, "False": False}[value]
    actual = context.data_label.number_format_is_linked
    assert actual is expected, "data_label.number_format_is_linked is %s" % actual


@then(
    "the c:dLbl for the point has c:numFmt with formatCode '{format_code}' "
    "and sourceLinked '{source_linked}'"
)
def then_c_dLbl_has_c_numFmt_formatCode_sourceLinked(context, format_code, source_linked):
    ser = context.data_label._ser
    numFmts = ser.xpath("c:dLbls/c:dLbl/c:numFmt")
    assert len(numFmts) == 1, "expected a single c:dLbl/c:numFmt, got %d" % len(numFmts)
    numFmt = numFmts[0]
    actual_format_code = numFmt.get("formatCode")
    assert actual_format_code == format_code, "c:numFmt/@formatCode is %r, expected %r" % (
        actual_format_code,
        format_code,
    )
    actual_source_linked = numFmt.get("sourceLinked")
    assert actual_source_linked == source_linked, "c:numFmt/@sourceLinked is %r, expected %r" % (
        actual_source_linked,
        source_linked,
    )


@then("data_label.position is {value}")
def then_data_label_position_is_value(context, value):
    expected_value = None if value == "None" else getattr(XL_DATA_LABEL_POSITION, value)
    data_label = context.data_label
    assert data_label.position is expected_value, "got %s" % data_label.position


@then("data_label.text_frame is a TextFrame object")
def then_data_label_text_frame_is_a_TextFrame_object(context):
    text_frame = context.data_label.text_frame
    assert type(text_frame).__name__ == "TextFrame"


@then("data_labels.position is {value}")
def then_data_labels_position_is_value(context, value):
    expected_value = None if value == "None" else getattr(XL_DATA_LABEL_POSITION, value)
    data_labels = context.data_labels
    assert data_labels.position is expected_value, "got %s" % data_labels.position


@then("data_labels.format is a ChartFormat object")
def then_data_labels_format_is_a_ChartFormat_object(context):
    data_labels = context.data_labels
    assert type(data_labels.format).__name__ == "ChartFormat"


@then("data_labels.format.fill is a FillFormat object")
def then_data_labels_format_fill_is_a_FillFormat_object(context):
    data_labels = context.data_labels
    assert type(data_labels.format.fill).__name__ == "FillFormat"


@then("data_labels.format.line is a LineFormat object")
def then_data_labels_format_line_is_a_LineFormat_object(context):
    data_labels = context.data_labels
    assert type(data_labels.format.line).__name__ == "LineFormat"


@then("data_labels.show_category_name is {value}")
def then_data_labels_show_category_name_is_value(context, value):
    actual, expected = context.data_labels.show_category_name, eval(value)
    assert actual is expected, "data_labels.show_category_name is %s" % actual


@then("data_labels.show_legend_key is {value}")
def then_data_labels_show_legend_key_is_value(context, value):
    actual, expected = context.data_labels.show_legend_key, eval(value)
    assert actual is expected, "data_labels.show_legend_key is %s" % actual


@then("data_labels.show_percentage is {value}")
def then_data_labels_show_percentage_is_value(context, value):
    actual, expected = context.data_labels.show_percentage, eval(value)
    assert actual is expected, "data_labels.show_percentage is %s" % actual


@then("data_labels.show_series_name is {value}")
def then_data_labels_show_series_name_is_value(context, value):
    actual, expected = context.data_labels.show_series_name, eval(value)
    assert actual is expected, "data_labels.show_series_name is %s" % actual


@then("data_labels.show_value is {value}")
def then_data_labels_show_value_is_value(context, value):
    actual, expected = context.data_labels.show_value, eval(value)
    assert actual is expected, "data_labels.show_value is %s" % actual


@then("data_labels.text_frame is a TextFrame object")
def then_data_labels_text_frame_is_a_TextFrame_object(context):
    text_frame = context.data_labels.text_frame
    assert type(text_frame).__name__ == "TextFrame"


@then("data_labels.text_frame wraps the c:dLbls/c:txPr element")
def then_data_labels_text_frame_wraps_the_c_dLbls_c_txPr_element(context):
    text_frame = context.data_labels.text_frame
    dLbls = context.data_labels._element
    txPrs = dLbls.xpath("c:txPr")
    assert len(txPrs) == 1, "expected a single c:txPr child, got %d" % len(txPrs)
    assert text_frame._txBody is txPrs[0], "TextFrame is not wrapping c:dLbls/c:txPr"


@then("data_labels.text_frame.word_wrap is {value}")
def then_data_labels_text_frame_word_wrap_is_value(context, value):
    expected = {"True": True, "False": False, "None": None}[value]
    actual = context.data_labels.text_frame.word_wrap
    assert actual is expected, "data_labels.text_frame.word_wrap is %s" % actual


@then('the c:dLbls XML has c:txPr/a:bodyPr with wrap="none"')
def then_the_c_dLbls_XML_has_c_txPr_a_bodyPr_with_wrap_none(context):
    dLbls = context.data_labels._element
    bodyPrs = dLbls.xpath("c:txPr/a:bodyPr")
    assert len(bodyPrs) == 1, "expected a single c:txPr/a:bodyPr, got %d" % len(bodyPrs)
    wrap = bodyPrs[0].get("wrap")
    assert wrap == "none", 'expected wrap="none", got wrap=%r' % wrap


@then("the c:dLbls XML has no c:tx/c:rich subtree")
def then_the_c_dLbls_XML_has_no_c_tx_c_rich_subtree(context):
    dLbls = context.data_labels._element
    assert dLbls.xpath("c:tx/c:rich") == [], "c:dLbls unexpectedly has c:tx/c:rich subtree"


@then("data_label.manual_layout is None")
def then_data_label_manual_layout_is_None(context):
    actual = context.data_label.manual_layout
    assert actual is None, "data_label.manual_layout is %r, expected None" % (actual,)


@then("data_label.manual_layout is ({x:g}, {y:g})")
def then_data_label_manual_layout_is_xy(context, x, y):
    from pptx.chart.datalabel import ManualLayout

    actual = context.data_label.manual_layout
    expected = ManualLayout(x, y)
    assert actual == expected, "data_label.manual_layout is %r, expected %r" % (actual, expected)
    assert isinstance(actual, ManualLayout), "expected ManualLayout, got %s" % type(actual).__name__


@then(
    "the c:dLbl for the point has c:layout/c:manualLayout with x={x:g} and y={y:g}",
)
def then_c_dLbl_has_manualLayout(context, x, y):
    ser = context.data_label._ser
    xs = ser.xpath("c:dLbls/c:dLbl/c:layout/c:manualLayout/c:x/@val")
    ys = ser.xpath("c:dLbls/c:dLbl/c:layout/c:manualLayout/c:y/@val")
    assert len(xs) == 1 and len(ys) == 1, (
        "expected a single c:manualLayout/c:x and c:y on the c:dLbl, got %d and %d"
        % (len(xs), len(ys))
    )
    assert float(xs[0]) == x, "c:manualLayout/c:x/@val is %r, expected %r" % (xs[0], x)
    assert float(ys[0]) == y, "c:manualLayout/c:y/@val is %r, expected %r" % (ys[0], y)
    # -- xMode/yMode must be "factor" (default) for the position to apply --
    xModes = ser.xpath("c:dLbls/c:dLbl/c:layout/c:manualLayout/c:xMode/@val")
    yModes = ser.xpath("c:dLbls/c:dLbl/c:layout/c:manualLayout/c:yMode/@val")
    for mode_val in xModes + yModes:
        assert mode_val == "factor", "c:xMode/c:yMode must be 'factor', got %r" % mode_val


@then("the c:dLbl for the point has no c:layout subtree")
def then_c_dLbl_has_no_c_layout(context):
    ser = context.data_label._ser
    layouts = ser.xpath("c:dLbls/c:dLbl/c:layout")
    assert layouts == [], "c:dLbl unexpectedly has c:layout subtree: %r" % layouts
