# encoding: utf-8

"""
Step implementations for line format features
"""

from __future__ import absolute_import, print_function, unicode_literals

from behave import given, then

from pptx import Presentation
from pptx.enum.dml import MSO_FILL_TYPE

from .helpers import cls_qname, test_pptx


# given ===================================================

@given('an autoshape having {outline_type}')
def given_autoshape_having_outline_type(context, outline_type):
    shape_idx = {
        'an inherited outline format': 0,
        'no outline':                  1,
        'a solid outline':             2,
    }[outline_type]
    prs = Presentation(test_pptx('shp-line-props'))
    autoshape = prs.slides[0].shapes[shape_idx]
    context.autoshape = autoshape


# then ====================================================

@then('I can access the line format of the shape')
def then_I_can_access_the_line_format_of_the_shape(context):
    shape = context.shape
    line_format = shape.line
    line_format_cls_name = cls_qname(line_format)
    expected_cls_name = 'pptx.dml.line.LineFormat'
    assert line_format_cls_name == expected_cls_name, (
        "expected '%s', got '%s'" % (expected_cls_name, line_format_cls_name)
    )


@then('the line fill type is {fill_type}')
def then_the_line_file_type_is_fill_type(context, fill_type):
    expected_fill_type = {
        'None':                      None,
        'MSO_FILL_TYPE.BACKGROUND':  MSO_FILL_TYPE.BACKGROUND,
        'MSO_FILL_TYPE.SOLID':       MSO_FILL_TYPE.SOLID,
    }[fill_type]
    shape = context.autoshape
    fill_type = shape.line.fill.type
    assert fill_type == expected_fill_type, (
        "expected '%s', got '%s'" % (expected_fill_type, fill_type)
    )
