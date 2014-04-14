# encoding: utf-8

"""
Step implementations for line format features
"""

from __future__ import absolute_import, print_function, unicode_literals

from behave import given, then, when

from pptx import Presentation
from pptx.enum.dml import MSO_FILL_TYPE

from .helpers import test_pptx


# given ===================================================

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
