# encoding: utf-8

"""
Gherkin step implementations for shape-related features.
"""

from __future__ import absolute_import, print_function

from behave import given, when, then

from pptx import Presentation
from pptx.chart.chart import Chart
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.action import ActionSetting
from pptx.util import Inches

from helpers import cls_qname, saved_pptx_path, test_pptx, test_text


# given ===================================================

@given('a chart')
def given_a_chart(context):
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[0]
    context.shape = sld.shapes[6]


@given('a connector')
def given_a_connector(context):
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[0]
    context.shape = sld.shapes[4]


@given('a graphic frame')  # shouldn't matter, but this one contains a table
def given_a_graphic_frame(context):
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[0]
    context.shape = sld.shapes[2]


@given('a graphic frame containing a chart')
def given_a_graphic_frame_containing_a_chart(context):
    prs = Presentation(test_pptx('shp-access-chart'))
    sld = prs.slides[0]
    context.shape = sld.shapes[0]


@given('a graphic frame containing a table')
def given_a_graphic_frame_containing_a_table(context):
    prs = Presentation(test_pptx('shp-access-chart'))
    sld = prs.slides[1]
    context.shape = sld.shapes[0]


@given('a group shape')
def given_a_group_shape(context):
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[0]
    context.shape = sld.shapes[3]


@given('a picture')
def given_a_picture(context):
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[0]
    context.shape = sld.shapes[1]


@given('a rotated {shape_type}')
def given_a_rotated_shape_type(context, shape_type):
    shape_idx = {
        'shape':         0,
        'picture':       1,
        'graphic frame': 2,
        'group shape':   3,
        'connector':     4,
    }[shape_type]
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[1]
    context.shape = sld.shapes[shape_idx]


@given('a shape')
def given_a_shape(context):
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[0]
    context.shape = sld.shapes[0]


@given('a table')
def given_a_table(context):
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[0]
    context.shape = sld.shapes[2]


@given('a textbox')
def given_a_textbox(context):
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[0]
    context.shape = sld.shapes[5]


@given('a {shape_type} on a slide')
def given_a_shape_on_a_slide(context, shape_type):
    shape_idx = {
        'shape':         0,
        'picture':       1,
        'graphic frame': 2,
        'group shape':   3,
        'connector':     4,
    }[shape_type]
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[0]
    context.shape = sld.shapes[shape_idx]
    context.slide = sld


@given('a shape of known position and size')
def given_a_shape_of_known_position_and_size(context):
    prs = Presentation(test_pptx('shp-pos-and-size'))
    context.shape = prs.slides[0].shapes[0]


@given('an autoshape')
def given_an_autoshape(context):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    shapes = prs.slides.add_slide(blank_slide_layout).shapes
    x = y = cx = cy = 914400
    context.shape = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, cx, cy)


@given('an autoshape having text')
def given_an_autoshape_having_text(context):
    prs = Presentation(test_pptx('shp-autoshape-props'))
    context.shape = prs.slides[0].shapes[0]


@given('I have a reference to a chevron shape')
def given_ref_to_chevron_shape(context):
    context.prs = Presentation()
    blank_slide_layout = context.prs.slide_layouts[6]
    shapes = context.prs.slides.add_slide(blank_slide_layout).shapes
    x = y = cx = cy = 914400
    context.chevron_shape = shapes.add_shape(MSO_SHAPE.CHEVRON, x, y, cx, cy)


# when ====================================================

@when("I add a text box to the slide's shape collection")
def when_I_add_a_text_box(context):
    shapes = context.slide.shapes
    x, y = (Inches(1.00), Inches(2.00))
    cx, cy = (Inches(3.00), Inches(1.00))
    sp = shapes.add_textbox(x, y, cx, cy)
    sp.text = test_text


@when("I add an auto shape to the slide's shape collection")
def when_I_add_an_auto_shape(context):
    shapes = context.slide.shapes
    x, y = (Inches(1.00), Inches(2.00))
    cx, cy = (Inches(3.00), Inches(4.00))
    sp = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, cx, cy)
    sp.text = test_text


@when("I assign a string to shape.text")
def when_I_assign_a_string_to_shape_text(context):
    context.shape.text = u'F\xf8o\nBar'


@when("I assign '{value}' to shape.name")
def when_I_assign_value_to_shape_name(context, value):
    context.shape.name = value


@when("I assign {value} to shape.rotation")
def when_I_assign_value_to_shape_rotation(context, value):
    context.shape.rotation = float(value)


@when("I change the left and top of the {shape_type}")
def when_I_change_the_position_of_the_shape(context, shape_type):
    left, top = {
        'shape':          (692696, 1339552),
        'picture':       (1835696, 2711152),
        'graphic frame': (2978696, 4082752),
        'group shape':   (4121696, 5454352),
        'connector':     (5264696, 6825952),
    }[shape_type]
    shape = context.shape
    shape.left = left
    shape.top = top


@when("I change the width and height of the {shape_type}")
def when_I_change_the_size_of_the_shape(context, shape_type):
    width, height = {
        'shape':          (692696, 1339552),
        'picture':       (1835696, 2711152),
        'graphic frame': (2978696, 4082752),
        'group shape':   (4121696, 5454352),
        'connector':     (5264696, 6825952),
    }[shape_type]
    shape = context.shape
    shape.width = width
    shape.height = height


@when("I get the chart from its graphic frame")
def when_I_get_the_chart_from_its_graphic_frame(context):
    context.chart = context.shape.chart


@when("I set the fill type to background")
def when_set_fill_type_to_background(context):
    context.shape.fill.background()


@when("I set the fill type to solid")
def when_set_fill_type_to_solid(context):
    context.shape.fill.solid()


@when("I set the first adjustment value to 0.15")
def when_set_first_adjustment_value(context):
    context.chevron_shape.adjustments[0] = 0.15


@when("I set the foreground color brightness to 0.5")
def when_set_fore_color_brightness_to_value(context):
    context.shape.fill.fore_color.brightness = 0.5


@when("I set the foreground color to a theme color")
def when_set_fore_color_to_theme_color(context):
    context.shape.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_6


@when("I set the foreground color to an RGB value")
def when_set_fore_color_to_RGB_value(context):
    context.shape.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)


# then ====================================================

@then('the auto shape appears in the slide')
def then_auto_shape_appears_in_slide(context):
    prs = Presentation(saved_pptx_path)
    sp = prs.slides[0].shapes[0]
    sp_text = sp.text_frame.paragraphs[0].runs[0].text
    assert sp.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
    assert sp.auto_shape_type == MSO_SHAPE.ROUNDED_RECTANGLE
    assert sp_text == test_text


@then('the chevron shape appears with a less acute arrow head')
def then_chevron_shape_appears_with_less_acute_arrow_head(context):
    chevron = Presentation(saved_pptx_path).slides[0].shapes[0]
    assert chevron.adjustments[0] == 0.15


@then('the fill type of the shape is background')
def then_fill_type_is_background(context):
    assert context.shape.fill.type == MSO_FILL.BACKGROUND


@then('the foreground color brightness of the shape is 0.5')
def then_fore_color_brightness_is_value(context):
    assert context.shape.fill.fore_color.brightness == 0.5


@then('the foreground color of the shape is the RGB value I set')
def then_fore_color_is_RGB_value_I_set(context):
    assert context.shape.fill.fore_color.rgb == RGBColor(0x12, 0x34, 0x56)


@then('the foreground color of the shape is the theme color I set')
def then_fore_color_is_theme_color_I_set(context):
    fore_color = context.shape.fill.fore_color
    assert fore_color.theme_color == MSO_THEME_COLOR.ACCENT_6


@then('I can access the line format of the shape')
def then_I_can_access_the_line_format_of_the_shape(context):
    shape = context.shape
    line_format = shape.line
    line_format_cls_name = cls_qname(line_format)
    expected_cls_name = 'pptx.dml.line.LineFormat'
    assert line_format_cls_name == expected_cls_name, (
        "expected '%s', got '%s'" % (expected_cls_name, line_format_cls_name)
    )


@then('I can access the slide from the shape')
def then_I_can_access_the_slide_from_the_shape(context):
    assert context.shape.part is context.slide


@then('I can determine the shape {has_text_frame_status}')
def then_the_shape_has_text_frame_status(context, has_text_frame_status):
    has_text_frame = {
        'has a text frame':  True,
        'has no text frame': False,
    }[has_text_frame_status]
    assert context.shape.has_text_frame is has_text_frame


@then('I can get the id of the {shape_type}')
def then_I_can_get_the_id_of_the_shape(context, shape_type):
    expected_id = {
        'shape':         2,
        'picture':       3,
        'graphic frame': 4,
        'group shape':   9,
        'connector':    11,
    }[shape_type]
    assert context.shape.id == expected_id


@then("shape.click_action is an ActionSetting object")
def then_shape_click_action_is_an_ActionSetting_object(context):
    assert isinstance(context.shape.click_action, ActionSetting)


@then("shape.name is '{expected_value}'")
def then_shape_name_is_value(context, expected_value):
    shape = context.shape
    msg = "expected shape name '%s', got '%s'" % (shape.name, expected_value)
    assert shape.name == expected_value, msg


@then("shape.rotation is {value}")
def then_shape_rotation_is_value(context, value):
    shape = context.shape
    expected_value = float(value)
    assert shape.rotation == expected_value, 'got %s' % expected_value


@then('shape.text is the string I assigned')
def then_shape_text_is_the_string_I_assigned(context):
    shape = context.shape
    assert shape.text == u'F\xf8o\nBar'


@then('shape.text is the text in the shape')
def then_shape_text_is_the_text_in_the_shape(context):
    shape = context.shape
    assert shape.text == u'Fee Fi\nF\xf8\xf8 Fum\nI am a shape\nwith textium'


@then('the chart is a Chart object')
def then_the_chart_is_a_Chart_object(context):
    assert isinstance(context.chart, Chart)


@then('the left and top of the {shape_type} match their new values')
def then_left_and_top_of_shape_match_new_values(context, shape_type):
    expected_left, expected_top = {
        'shape':          (692696, 1339552),
        'picture':       (1835696, 2711152),
        'graphic frame': (2978696, 4082752),
        'group shape':   (4121696, 5454352),
        'connector':     (5264696, 6825952),
    }[shape_type]
    shape = context.shape
    assert shape.left == expected_left, 'got left: %s' % shape.left
    assert shape.top == expected_top, 'got top: %s' % shape.top


@then('the left and top of the {shape_type} match their known values')
def then_left_and_top_of_shape_match_known_values(context, shape_type):
    expected_left, expected_top = {
        'shape':         (1339552,  692696),
        'picture':       (2711152, 1835696),
        'graphic frame': (4082752, 2978696),
        'group shape':   (5454352, 4121696),
        'connector':     (6825952, 5264696),
    }[shape_type]
    shape = context.shape
    assert shape.left == expected_left, 'got left: %s' % shape.left
    assert shape.top == expected_top, 'got top: %s' % shape.top


@then('the shape {has_or_not} a chart')
def then_the_shape_has_or_not_a_chart(context, has_or_not):
    expected_bool = {'has': True, 'does not have': False}[has_or_not]
    shape = context.shape
    assert shape.has_chart is expected_bool


@then('the width and height of the {shape_type} match their known values')
def then_width_and_height_of_shape_match_known_values(context, shape_type):
    expected_width, expected_height = {
        'shape':         (928192, 914400),
        'picture':       (914400, 945232),
        'graphic frame': (993304, 914400),
        'group shape':   (914400, 914400),
        'connector':     (986408, 828600),
    }[shape_type]
    shape = context.shape
    assert shape.width == expected_width, 'got width: %s' % shape.width
    assert shape.height == expected_height, 'got height: %s' % shape.height


@then('the width and height of the {shape_type} match their new values')
def then_width_and_height_of_shape_match_new_values(context, shape_type):
    expected_width, expected_height = {
        'shape':          (692696, 1339552),
        'picture':       (1835696, 2711152),
        'graphic frame': (2978696, 4082752),
        'group shape':   (4121696, 5454352),
        'connector':     (5264696, 6825952),
    }[shape_type]
    shape = context.shape
    assert shape.width == expected_width, 'got width: %s' % shape.width
    assert shape.height == expected_height, 'got height: %s' % shape.height


@then('the text box appears in the slide')
def then_text_box_appears_in_slide(context):
    prs = Presentation(saved_pptx_path)
    textbox = prs.slides[0].shapes[0]
    textbox_text = textbox.text_frame.paragraphs[0].runs[0].text
    assert textbox_text == test_text
