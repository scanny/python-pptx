# encoding: utf-8

"""
Gherkin step implementations for table-related features.
"""

from __future__ import absolute_import

from behave import given, when, then

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches

from helpers import saved_pptx_path, test_pptx


# given ===================================================


@given('a 2x2 table')
def given_a_2x2_table(context):
    prs = Presentation(test_pptx('shp-shape-access'))
    context.prs = prs
    context.table_ = prs.slides[0].shapes[3].table


@given('a table cell')
def given_a_table_cell(context):
    prs = Presentation(test_pptx('shp-shape-access'))
    context.prs = prs
    context.cell = prs.slides[0].shapes[3].table.cell(0, 0)


# when ====================================================


@when("I add a table to the slide's shape collection")
def when_I_add_a_table(context):
    shapes = context.slide.shapes
    x, y = (Inches(1.00), Inches(2.00))
    cx, cy = (Inches(3.00), Inches(1.00))
    shapes.add_table(2, 2, x, y, cx, cy)


@when("I set the cell fill foreground color to an RGB value")
def when_set_cell_fore_color_to_RGB_value(context):
    context.cell.fill.fore_color.rgb = RGBColor(0x23, 0x45, 0x67)


@when("I set the cell fill type to solid")
def when_set_cell_fill_type_to_solid(context):
    context.cell.fill.solid()


@when("I set the cell margins")
def when_set_cell_margins(context):
    context.cell.margin_top = 1000
    context.cell.margin_right = 2000
    context.cell.margin_bottom = 3000
    context.cell.margin_left = 4000


@when("I set the cell vertical anchor to middle")
def when_set_cell_vertical_anchor_to_middle(context):
    context.cell.vertical_anchor = MSO_ANCHOR.MIDDLE


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


@when("I set the text of the first cell")
def when_set_text_of_first_cell(context):
    context.table_.cell(0, 0).text = 'test text'


@when("I set the vert_banding property to True")
def when_set_vert_banding_property_to_true(context):
    context.table_.vert_banding = True


@when("I set the width of the table's columns")
def when_set_table_column_widths(context):
    context.table_.columns[0].width = Inches(1.50)
    context.table_.columns[1].width = Inches(3.00)


# then ====================================================


@then('the cell contents are inset by the margins')
def then_cell_contents_are_inset_by_the_margins(context):
    prs = Presentation(saved_pptx_path)
    table = prs.slides[0].shapes[3].table
    cell = table.cell(0, 0)
    assert cell.margin_top == 1000
    assert cell.margin_right == 2000
    assert cell.margin_bottom == 3000
    assert cell.margin_left == 4000


@then('the cell contents are vertically centered')
def then_cell_contents_are_vertically_centered(context):
    prs = Presentation(saved_pptx_path)
    table = prs.slides[0].shapes[3].table
    cell = table.cell(0, 0)
    assert cell.vertical_anchor == MSO_ANCHOR.MIDDLE


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


@then('the foreground color of the cell is the RGB value I set')
def then_cell_fore_color_is_RGB_value_I_set(context):
    assert context.cell.fill.fore_color.rgb == RGBColor(0x23, 0x45, 0x67)


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


@then('the text appears in the first cell of the table')
def then_text_appears_in_first_cell_of_table(context):
    prs = Presentation(saved_pptx_path)
    table = prs.slides[0].shapes[3].table
    text = table.cell(0, 0).text_frame.paragraphs[0].runs[0].text
    assert text == 'test text'
