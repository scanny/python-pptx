# encoding: utf-8

"""
Step implementations for line format features
"""

from __future__ import absolute_import, print_function, unicode_literals

from behave import given, then, when

from pptx import Presentation
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.util import Length, Pt

from helpers import test_pptx


# given ===================================================

@given('a line of {line_width} width')
def given_a_line_of_width(context, line_width):
    shape_idx = {
        'no explicit': 0,
        '1 pt':        1,
    }[line_width]
    prs = Presentation(test_pptx('dml-line'))
    shape = prs.slides[2].shapes[shape_idx]
    context.line = shape.line


@given('a LineFormat object as line')
def given_a_LineFormat_object_as_line(context):
    line = Presentation(test_pptx('dml-line')).slides[0].shapes[0].line
    context.line = line


@given('an autoshape outline having {outline_type}')
def given_autoshape_outline_having_outline_type(context, outline_type):
    shape_idx = {
        'an inherited outline format': 0,
        'no outline':                  1,
        'a solid outline':             2,
    }[outline_type]
    prs = Presentation(test_pptx('dml-line'))
    autoshape = prs.slides[0].shapes[shape_idx]
    context.line = autoshape.line


# when ====================================================

@when('I set the line width to {line_width}')
def when_I_set_the_line_width_to_value(context, line_width):
    value = {
        'None':    None,
        '1 pt':    Pt(1),
        '2.34 pt': Pt(2.34),
    }[line_width]
    context.line.width = value


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


@then('line.color is a ColorFormat object')
def then_line_color_is_a_ColorFormat_object(context):
    class_name = context.line.color.__class__.__name__
    expected_value = 'ColorFormat'
    assert class_name == expected_value, (
        'expected \'%s\', got \'%s\'' % (expected_value, class_name)
    )


@then("the reported line width is {line_width}")
def then_the_reported_line_width_is_value(context, line_width):
    expected_value = {
        '0':       0,
        '1 pt':    Pt(1),
        '2.34 pt': Pt(2.34),
    }[line_width]
    line_width = context.line.width
    assert line_width == expected_value
    assert isinstance(line_width, Length)
