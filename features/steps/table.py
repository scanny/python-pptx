# encoding: utf-8

"""Gherkin step implementations for table-related features"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from behave import given, when, then

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR  # noqa
from pptx.util import Inches

from helpers import saved_pptx_path, test_pptx


# given ===================================================

@given('a 2x2 table')
def given_a_2x2_table(context):
    prs = Presentation(test_pptx('shp-shapes'))
    context.prs = prs
    context.table_ = prs.slides[0].shapes[3].table


@given('a _Cell object as cell')
def given_a_Cell_object_as_cell(context):
    prs = Presentation(test_pptx('shp-shapes'))
    context.cell = prs.slides[0].shapes[3].table.cell(0, 0)


@given('a _Cell object with known margins as cell')
def given_a_Cell_object_with_known_margins_as_cell(context):
    prs = Presentation(test_pptx('tbl-cell'))
    context.cell = prs.slides[0].shapes[0].table.cell(0, 0)


@given('a _Cell object with {setting} vertical alignment as cell')
def given_a_Cell_object_with_setting_vertical_alignment(context, setting):
    cell_coordinates = {
        'inherited': (0, 1),
        'middle': (0, 2),
        'bottom': (0, 3),
    }[setting]
    prs = Presentation(test_pptx('tbl-cell'))
    context.cell = prs.slides[0].shapes[0].table.cell(*cell_coordinates)


# when ====================================================

@when("I add a table to the slide's shape collection")
def when_I_add_a_table(context):
    shapes = context.slide.shapes
    x, y = (Inches(1.00), Inches(2.00))
    cx, cy = (Inches(3.00), Inches(1.00))
    shapes.add_table(2, 2, x, y, cx, cy)


@when('I assign cell.margin_{side} = {value}')
def when_I_assign_cell_margin_side_eq_value(context, value, side):
    setattr(context.cell, 'margin_%s' % side, eval(value))


@when('I assign cell.text = "test text"')
def when_I_assign_cell_text(context):
    context.cell.text = 'test text'


@when('I assign cell.vertical_anchor = {value}')
def when_I_assign_cell_vertical_anchor_eq_value(context, value):
    context.cell.vertical_anchor = eval(value)


@when("I set the first_col property to True")
def when_set_first_col_property_to_true(context):
    context.table_.first_col = True


@when("I set the first_row property to True")
def when_set_first_row_property_to_true(context):
    context.table_.first_row = True


@when("I set the horz_banding property to True")
def when_set_horz_banding_property_to_true(context):
    context.table_.horz_banding = True


@when("I set the last_col property to True")
def when_set_last_col_property_to_true(context):
    context.table_.last_col = True


@when("I set the last_row property to True")
def when_set_last_row_property_to_true(context):
    context.table_.last_row = True


@when("I set the vert_banding property to True")
def when_set_vert_banding_property_to_true(context):
    context.table_.vert_banding = True


@when("I set the width of the table's columns")
def when_set_table_column_widths(context):
    context.table_.columns[0].width = Inches(1.50)
    context.table_.columns[1].width = Inches(3.00)


# then ====================================================

@then('cell.fill is a FillFormat object')
def then_cell_fill_is_a_FillFormat_object(context):
    actual = type(context.cell.fill).__name__
    expected = 'FillFormat'
    assert actual == expected, 'cell.fill is a %s object' % actual


@then('cell.margin_{side} == Inches({num_lit})')
def then_cell_margin_side_eq_Inches_num(context, side, num_lit):
    actual = getattr(context.cell, 'margin_%s' % side)
    expected = Inches(float(num_lit))
    assert actual == expected, 'cell.margin_%s == %s' % (side, actual.inches)


@then('cell.text_frame.text == "test text"')
def then_cell_text_frame_text_eq_test_text(context):
    actual = context.cell.text_frame.text
    expected = 'test text'
    assert actual == expected, 'cell.text_frame.text == %s' % actual


@then('cell.vertical_anchor == {value}')
def then_cell_vertical_anchor_eq_value(context, value):
    actual = context.cell.vertical_anchor
    expected = eval(value)
    assert actual == expected, 'cell.vertical_anchor == %s' % actual


@then('the columns of the table have alternating shading')
def then_columns_of_table_have_alternating_shading(context):
    table = Presentation(saved_pptx_path).slides[0].shapes[3].table
    assert table.vert_banding is True


@then('the first column of the table has special formatting')
def then_first_column_of_table_has_special_formatting(context):
    table = Presentation(saved_pptx_path).slides[0].shapes[3].table
    assert table.first_col is True


@then('the first row of the table has special formatting')
def then_first_row_of_table_has_special_formatting(context):
    table = Presentation(saved_pptx_path).slides[0].shapes[3].table
    assert table.first_row is True


@then('the last column of the table has special formatting')
def then_last_column_of_table_has_special_formatting(context):
    table = Presentation(saved_pptx_path).slides[0].shapes[3].table
    assert table.last_col is True


@then('the last row of the table has special formatting')
def then_last_row_of_table_has_special_formatting(context):
    table = Presentation(saved_pptx_path).slides[0].shapes[3].table
    assert table.last_row is True


@then('the rows of the table have alternating shading')
def then_rows_of_table_have_alternating_shading(context):
    table = Presentation(saved_pptx_path).slides[0].shapes[3].table
    assert table.horz_banding is True


@then('the table appears in the slide')
def then_the_table_appears_in_the_slide(context):
    prs = Presentation(saved_pptx_path)
    expected_table_graphic_frame = prs.slides[0].shapes[0]
    assert expected_table_graphic_frame.has_table


@then('the table appears with the new column widths')
def then_table_appears_with_new_col_widths(context):
    prs = Presentation(saved_pptx_path)
    table = prs.slides[0].shapes[3].table
    assert table.columns[0].width == Inches(1.50)
    assert table.columns[1].width == Inches(3.00)
