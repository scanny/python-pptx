# encoding: utf-8

"""
Step implementations for line format features
"""

from __future__ import absolute_import, print_function, unicode_literals

from behave import given, then, when

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL_TYPE, MSO_THEME_COLOR

from .helpers import test_pptx


# given ===================================================

@given('a line with {color_type} color')
def given_a_line_with_color_type_color(context, color_type):
    shape_idx = {
        'no':      0,
        'an RGB':  1,
        'a theme': 2
    }[color_type]
    prs = Presentation(test_pptx('shp-line-props'))
    shape = prs.slides[1].shapes[shape_idx]
    context.line = shape.line


@given('an autoshape outline having {outline_type}')
def given_autoshape_outline_having_outline_type(context, outline_type):
    shape_idx = {
        'an inherited outline format': 0,
        'no outline':                  1,
        'a solid outline':             2,
    }[outline_type]
    prs = Presentation(test_pptx('shp-line-props'))
    autoshape = prs.slides[0].shapes[shape_idx]
    context.line = autoshape.line


# when ====================================================

@when('I set the line {color_type} value')
def when_I_set_the_line_color_value(context, color_type):
    if color_type == 'RGB':
        context.line.color.rgb = RGBColor(0x12, 0x34, 0x56)
    elif color_type == 'theme color':
        context.line.color.theme_color = MSO_THEME_COLOR.DARK_1


@when('I set the line fill type to {line_fill_type}')
def when_I_set_the_line_fill_type_to_line_fill_type(context, line_fill_type):
    line = context.line
    if line_fill_type == 'solid':
        line.fill.solid()
    elif line_fill_type == 'background':
        line.fill.background()


# then ====================================================

@then('the line fill type is {fill_type}')
def then_the_line_file_type_is_fill_type(context, fill_type):
    expected_fill_type = {
        'None':                      None,
        'MSO_FILL_TYPE.BACKGROUND':  MSO_FILL_TYPE.BACKGROUND,
        'MSO_FILL_TYPE.SOLID':       MSO_FILL_TYPE.SOLID,
    }[fill_type]
    line = context.line
    fill_type = line.fill.type
    assert fill_type == expected_fill_type, (
        "expected '%s', got '%s'" % (expected_fill_type, fill_type)
    )


@then("the line's color type is {color_type}")
def then_the_line_color_type_is_value(context, color_type):
    expected_value = {
        'None':        None,
        'RGB':         MSO_COLOR_TYPE.RGB,
        'theme color': MSO_COLOR_TYPE.SCHEME,
    }[color_type]
    line = context.line
    assert line.color.type == expected_value


@then("the line's {color_type} value matches the new value")
def then_the_line_color_type_value_matches(context, color_type):
    line = context.line
    if color_type == 'RGB':
        assert line.color.rgb == RGBColor(0x12, 0x34, 0x56)
    else:
        assert line.color.theme_color == MSO_THEME_COLOR.DARK_1
