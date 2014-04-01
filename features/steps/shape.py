# encoding: utf-8

"""
Gherkin step implementations for shape-related features.
"""

from __future__ import absolute_import, print_function

from behave import given, when, then
from hamcrest import assert_that, equal_to, is_

from pptx import Presentation
from pptx.constants import MSO_AUTO_SHAPE_TYPE as MAST, MSO
from pptx.enum import MSO_FILL_TYPE as MSO_FILL, MSO_THEME_COLOR
from pptx.dml.color import RGBColor
from pptx.util import Inches

from .helpers import saved_pptx_path, test_pptx, test_text


# given ===================================================

@given('a connector')
def given_a_connector(context):
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[0]
    context.shape = sld.shapes[4]


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


@given('a {shape_type} on a slide')
def given_a_shape_on_a_slide(context, shape_type):
    shape_idx = {
        'shape':       0,
        'picture':     1,
        'table':       2,
        'group shape': 3,
        'connector':   4,
    }[shape_type]
    prs = Presentation(test_pptx('shp-common-props'))
    sld = prs.slides[0]
    context.shape = sld.shapes[shape_idx]
    context.slide = sld


@given('a shape of known position and size')
def given_a_shape_of_known_position_and_size(context):
    prs = Presentation(test_pptx('shp-pos-and-size'))
    context.shape = prs.slides[0].shapes[0]


@given('a table of known position and size')
def given_a_table_of_known_position_and_size(context):
    prs = Presentation(test_pptx('shp-pos-and-size'))
    context.shape = prs.slides[2].shapes[0]


@given('an autoshape')
def given_an_autoshape(context):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    shapes = prs.slides.add_slide(blank_slide_layout).shapes
    x = y = cx = cy = 914400
    context.shape = shapes.add_shape(MAST.ROUNDED_RECTANGLE, x, y, cx, cy)


@given('I have a reference to a chevron shape')
def given_ref_to_chevron_shape(context):
    context.prs = Presentation()
    blank_slide_layout = context.prs.slide_layouts[6]
    shapes = context.prs.slides.add_slide(blank_slide_layout).shapes
    x = y = cx = cy = 914400
    context.chevron_shape = shapes.add_shape(MAST.CHEVRON, x, y, cx, cy)


# when ====================================================

@when("I add a text box to the slide's shape collection")
def when_add_text_box(context):
    shapes = context.sld.shapes
    x, y = (Inches(1.00), Inches(2.00))
    cx, cy = (Inches(3.00), Inches(1.00))
    sp = shapes.add_textbox(x, y, cx, cy)
    sp.text = test_text


@when("I add an auto shape to the slide's shape collection")
def when_add_auto_shape(context):
    shapes = context.sld.shapes
    x, y = (Inches(1.00), Inches(2.00))
    cx, cy = (Inches(3.00), Inches(4.00))
    sp = shapes.add_shape(MAST.ROUNDED_RECTANGLE, x, y, cx, cy)
    sp.text = test_text


@when("I change the position of the connector")
def when_I_change_the_position_of_the_connector(context):
    shape = context.shape
    shape.left = 914400*4
    shape.top = 914400*3


@when("I change the position of the group shape")
def when_I_change_the_position_of_the_group_shape(context):
    shape = context.shape
    shape.left = 914400*4
    shape.top = 914400*3
    # shape.width = 914400*2
    # shape.height = 914400*1


@when("I change the position and size of the picture")
def when_change_pos_and_size_of_picture(context):
    picture = context.picture
    picture.left = 914400*4
    picture.top = 914400*3
    picture.width = 914400*2
    picture.height = 914400*1


@when("I change the position and size of the shape")
def when_change_pos_and_size_of_shape(context):
    shape = context.shape
    shape.left = 914400*4
    shape.top = 914400*3
    shape.width = 914400*2
    shape.height = 914400*1


@when("I change the position of the table")
def when_change_position_of_table(context):
    table = context.shape
    table.left = 914400*4
    table.top = 914400*3


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
    sp_text = sp.textframe.paragraphs[0].runs[0].text
    assert_that(sp.shape_type, is_(equal_to(MSO.AUTO_SHAPE)))
    assert_that(sp.auto_shape_type, is_(equal_to(MAST.ROUNDED_RECTANGLE)))
    assert_that(sp_text, is_(equal_to(test_text)))


@then('the chevron shape appears with a less acute arrow head')
def then_chevron_shape_appears_with_less_acute_arrow_head(context):
    chevron = Presentation(saved_pptx_path).slides[0].shapes[0]
    assert_that(chevron.adjustments[0], is_(equal_to(0.15)))


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


@then('I can access the slide from the shape')
def then_I_can_access_the_slide_from_the_shape(context):
    assert context.shape.part is context.slide


@then('I can determine the shape {has_textframe_status}')
def then_the_shape_has_textframe_status(context, has_textframe_status):
    has_textframe = {
        'has a text frame':  True,
        'has no text frame': False,
    }[has_textframe_status]
    assert context.shape.has_textframe is has_textframe


@then('I can get the id of the {shape_type}')
def then_I_can_get_the_id_of_the_shape(context, shape_type):
    expected_id = {
        'shape':        2,
        'picture':      3,
        'table':        4,
        'group shape':  9,
        'connector':   11,
    }[shape_type]
    assert context.shape.id == expected_id


@then('I can get the name of the {shape_type}')
def then_I_can_get_the_name_of_the_shape(context, shape_type):
    expected_name = {
        'shape':       'Rounded Rectangle 1',
        'picture':     'Picture 2',
        'table':       'Table 3',
        'group shape': 'Group 8',
        'connector':   'Elbow Connector 10',
    }[shape_type]
    shape = context.shape
    msg = "expected shape name '%s', got '%s'" % (shape.name, expected_name)
    assert shape.name == expected_name, msg


@then('the left and top of the connector match their known values')
def then_the_left_and_top_of_the_connector_match_known_values(context):
    shape = context.shape
    assert shape.left == 6825952, 'got %s' % shape.left
    assert shape.top == 5264696, 'got %s' % shape.top


@then('the left and top of the connector match the new values')
def then_the_left_and_top_of_the_connector_match_the_new_values(context):
    shape = context.shape
    assert shape.left == 914400*4
    assert shape.top == 914400*3


@then('the left and top of the group shape match their known values')
def then_the_left_and_top_of_the_group_shape_match_known_values(context):
    shape = context.shape
    assert shape.left == 5454352, 'got %s' % shape.left
    assert shape.top == 4121696, 'got %s' % shape.top
    # assert shape.width == 914400, 'got %s' % shape.width
    # assert shape.height == 914400, 'got %s' % shape.height


@then('the left and top of the group shape match the new values')
def then_the_left_and_top_of_the_group_shape_match_the_new_values(context):
    shape = context.shape
    assert shape.left == 914400*4
    assert shape.top == 914400*3
    # assert shape.width == 914400*2
    # assert shape.height == 914400*1


@then('the position and size of the picture matches the known values')
def then_picture_pos_and_size_matches_known_values(context):
    picture = context.picture
    assert picture.left == 914400
    assert picture.top == 914400*2
    assert picture.width == 914400*3
    assert picture.height == 914400*4


@then('the position and size of the picture matches the new values')
def then_picture_pos_and_size_matches_new_values(context):
    picture = context.picture
    assert picture.left == 914400*4
    assert picture.top == 914400*3
    assert picture.width == 914400*2
    assert picture.height == 914400*1


@then('the position and size of the shape matches the known values')
def then_shape_pos_and_size_matches_known_values(context):
    shape = context.shape
    assert shape.left == 914400
    assert shape.top == 914400*2
    assert shape.width == 914400*3
    assert shape.height == 914400*4


@then('the position and size of the shape matches the new values')
def then_shape_pos_and_size_matches_new_values(context):
    shape = context.shape
    assert shape.left == 914400*4
    assert shape.top == 914400*3
    assert shape.width == 914400*2
    assert shape.height == 914400*1


@then('the position and size of the table match the known values')
def then_table_pos_and_size_match_known_values(context):
    table = context.shape
    assert table.left == 1524000
    assert table.top == 1397000
    assert table.width == 6096000
    assert table.height == 741680


@then('the position of the table matches the new values')
def then_table_position_matches_the_new_values(context):
    table = context.shape
    assert table.left == 914400*4
    assert table.top == 914400*3


@then('the text box appears in the slide')
def then_text_box_appears_in_slide(context):
    prs = Presentation(saved_pptx_path)
    textbox = prs.slides[0].shapes[0]
    textbox_text = textbox.textframe.paragraphs[0].runs[0].text
    assert_that(textbox_text, is_(equal_to(test_text)))
