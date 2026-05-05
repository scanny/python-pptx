"""Gherkin step implementations for table-related features"""

from __future__ import annotations

from behave import given, then, when
from helpers import test_pptx

from pptx import Presentation
from pptx.dml.color import RGBColor  # noqa # pyright: ignore[reportUnusedImport]
from pptx.enum.dml import MSO_LINE  # noqa # pyright: ignore[reportUnusedImport]
from pptx.enum.text import MSO_ANCHOR  # noqa # pyright: ignore[reportUnusedImport]
from pptx.util import Inches, Pt  # noqa # pyright: ignore[reportUnusedImport]

# given ===================================================


@given("a Table object as table")
@given("a 2x2 Table object as table")
def given_a_2x2_Table_object_as_table(context):
    prs = Presentation(test_pptx("shp-shapes"))
    context.table_ = prs.slides[0].shapes[3].table


@given("a 2x3 _MergeOriginCell object as cell")
def given_a_2x3_MergeOriginCell_object_as_cell(context):
    prs = Presentation(test_pptx("tbl-cell"))
    context.cell = prs.slides[1].shapes[1].table.cell(0, 0)


@given("a 3x3 Table object as table")
@given("a 3x3 Table object with cells a to i as table")
def given_a_3x3_table_with_cells_a_to_i_as_table(context):
    prs = Presentation(test_pptx("tbl-cell"))
    # ---context.table is used by Behave for some odd reason---
    context.table_ = prs.slides[2].shapes[0].table


@given("a freshly-added {rows:d}x{cols:d} Table object as table")
def given_a_freshly_added_RxC_Table_object_as_table(context, rows: int, cols: int):
    """Create a fresh presentation + slide + table of the requested size.

    Used by the #636 scenarios where the table has to be a specific shape
    (e.g. 5x1 or 1x5) and the pre-canned fixture decks don't have a match.
    """
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    graphic_frame = slide.shapes.add_table(
        rows, cols, Inches(1), Inches(1), Inches(6), Inches(2)
    )
    context.table_ = graphic_frame.table


@given("a _Cell object as cell")
def given_a_Cell_object_as_cell(context):
    prs = Presentation(test_pptx("shp-shapes"))
    context.cell = prs.slides[0].shapes[3].table.cell(0, 0)


@given('a _Cell object containing "unladen swallows" as cell')
def given_a_Cell_object_containing_unladen_swallows_as_cell(context):
    prs = Presentation(test_pptx("tbl-cell"))
    context.cell = prs.slides[0].shapes[0].table.cell(1, 0)


@given("a _Cell object with known margins as cell")
def given_a_Cell_object_with_known_margins_as_cell(context):
    prs = Presentation(test_pptx("tbl-cell"))
    context.cell = prs.slides[0].shapes[0].table.cell(0, 0)


@given("a _Cell object with {setting} vertical alignment as cell")
def given_a_Cell_object_with_setting_vertical_alignment(context, setting):
    cell_coordinates = {"inherited": (0, 1), "middle": (0, 2), "bottom": (0, 3)}[setting]
    prs = Presentation(test_pptx("tbl-cell"))
    context.cell = prs.slides[0].shapes[0].table.cell(*cell_coordinates)


@given("a {role} _Cell object as cell")
def given_a_role_Cell_object_as_cell(context, role):
    coordinates = {"merge-origin": (0, 0), "spanned": (0, 1), "unmerged": (2, 2)}[role]
    table = Presentation(test_pptx("tbl-cell")).slides[1].shapes[0].table
    # ---create other_cell here where we know the coordinates---
    context.cell = table.cell(*coordinates)
    context.other_cell = table.cell(*coordinates)


@given("a second proxy instance for that cell as other_cell")
def given_a_second_proxy_instance_for_that_cell_as_other_cell(context):
    # ---other_cell is actually produced by prior step---
    assert context.other_cell


@given("a _Column object as column")
def given_a_Column_object_as_column(context):
    prs = Presentation(test_pptx("shp-shapes"))
    context.column = prs.slides[0].shapes[3].table.columns[0]


@given("a fully-styled source _Cell and a plain target _Cell")
def given_a_fully_styled_source_cell_and_a_plain_target_cell(context):
    """Build a freshly-added 2x2 table, fully style cell(0,0), leave cell(1,1) plain."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    gf = slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(6), Inches(2))
    table = gf.table

    source = table.cell(0, 0)
    source.text = "styled"
    source.fill.solid()
    source.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)
    source.border_left.color.rgb = RGBColor(0x00, 0x00, 0xFF)
    source.border_left.width = Pt(1.5)
    source.border_diagonal_down.color.rgb = RGBColor(0x00, 0xFF, 0x00)
    source.margin_left = Inches(0.2)
    source.vertical_anchor = MSO_ANCHOR.MIDDLE

    context.source_cell = source
    context.source_xml_snapshot = source._tc.xml  # pyright: ignore[reportPrivateUsage]
    context.target_cell = table.cell(1, 1)


# when ====================================================


@when("I assign cell.margin_{side} = {value}")
def when_I_assign_cell_margin_side_eq_value(context, value, side):
    setattr(context.cell, "margin_%s" % side, eval(value))


@when("I set cell.border_left to a red, 1.5pt, dashed line")
def when_I_set_cell_border_left_to_red_dashed(context):
    border = context.cell.border_left
    border.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    border.width = Pt(1.5)
    border.dash_style = MSO_LINE.DASH


@when('I assign cell.text = "test text"')
def when_I_assign_cell_text(context):
    context.cell.text = "test text"


@when("I assign cell.vertical_anchor = {value}")
def when_I_assign_cell_vertical_anchor_eq_value(context, value):
    context.cell.vertical_anchor = eval(value)


@when("I assign column.width = {value}")
def when_I_assign_column_width_eq_value(context, value):
    context.column.width = eval(value)


@when("I assign origin_cell = table.cell(0, 0)")
def when_I_assign_origin_cell_eq_table_cell_0_0(context):
    context.origin_cell = context.table_.cell(0, 0)


@when("I assign origin_cell = table.cell({row:d}, {col:d})")
def when_I_assign_origin_cell_eq_table_cell_r_c(context, row: int, col: int):
    context.origin_cell = context.table_.cell(row, col)


@when("I assign other_cell = table.cell(1, 1)")
def when_I_assign_other_cell_eq_table_cell_1_1(context):
    context.other_cell = context.table_.cell(1, 1)


@when("I assign other_cell = table.cell({row:d}, {col:d})")
def when_I_assign_other_cell_eq_table_cell_r_c(context, row: int, col: int):
    context.other_cell = context.table_.cell(row, col)


@when("I assign table.first_col = True")
def when_I_assign_table_first_col_eq_True(context):
    context.table_.first_col = True


@when("I assign table.first_row = True")
def when_I_assign_table_first_row_eq_True(context):
    context.table_.first_row = True


@when("I assign table.horz_banding = True")
def when_I_assign_table_horz_banding_eq_True(context):
    context.table_.horz_banding = True


@when("I assign table.last_col = True")
def when_I_assign_table_last_col_eq_True(context):
    context.table_.last_col = True


@when("I assign table.last_row = True")
def when_I_assign_table_last_row_eq_True(context):
    context.table_.last_row = True


@when("I assign table.style_id = {value}")
def when_I_assign_table_style_id_eq_value(context, value):
    context.table_.style_id = eval(value)


@when("I assign table.vert_banding = True")
def when_I_assign_table_vert_banding_eq_True(context):
    context.table_.vert_banding = True


@when("I call table.rows.add()")
def when_I_call_table_rows_add(context):
    # ---remember the prior last row's height before adding---
    context.prior_last_row_height = context.table_.rows[len(context.table_.rows) - 1].height
    context.added_row = context.table_.rows.add()


@when("I call table.add_row()")
def when_I_call_table_add_row(context):
    # ---remember the prior last row's height before adding---
    context.prior_last_row_height = context.table_.rows[len(context.table_.rows) - 1].height
    context.added_row = context.table_.add_row()


@when("I call table.rows[0].delete()")
def when_I_call_table_rows_0_delete(context):
    context.table_.rows[0].delete()


@when("I call table.columns.add()")
def when_I_call_table_columns_add(context):
    # ---remember the prior last column's width before adding---
    context.prior_last_column_width = context.table_.columns[
        len(context.table_.columns) - 1
    ].width
    context.added_column = context.table_.columns.add()


@when("I call table.add_column()")
def when_I_call_table_add_column(context):
    # ---remember the prior last column's width before adding---
    context.prior_last_column_width = context.table_.columns[
        len(context.table_.columns) - 1
    ].width
    context.added_column = context.table_.add_column()


@when("I call table.columns[0].delete()")
def when_I_call_table_columns_0_delete(context):
    context.table_.columns[0].delete()


@when("I call cell.split()")
def when_I_call_cell_split_other_cell(context):
    context.cell.split()


@when("I call origin_cell.merge(other_cell)")
def when_I_call_origin_cell_merge_other_cell(context):
    context.origin_cell.merge(context.other_cell)


@when("I call target_cell.clone_from(source_cell)")
def when_I_call_target_clone_from_source(context):
    context.clone_from_result = context.target_cell.clone_from(context.source_cell)


# then ====================================================


@then("cell == other_cell")
def then_cell_eq_other_cell(context):
    cell, other_cell = context.cell, context.other_cell
    assert cell == other_cell, "cell != other_cell"


@then("cell.fill is a FillFormat object")
def then_cell_fill_is_a_FillFormat_object(context):
    actual = type(context.cell.fill).__name__
    expected = "FillFormat"
    assert actual == expected, "cell.fill is a %s object" % actual


@then("cell.{border} is a LineFormat object")
def then_cell_border_is_a_LineFormat_object(context, border):
    actual = type(getattr(context.cell, border)).__name__
    expected = "LineFormat"
    assert actual == expected, "cell.%s is a %s object" % (border, actual)


@then("cell.border_left.color.rgb is RGBColor({r}, {g}, {b})")
def then_cell_border_left_color_rgb_eq(context, r, g, b):
    actual = context.cell.border_left.color.rgb
    expected = RGBColor(int(r, 16), int(g, 16), int(b, 16))
    assert actual == expected, "cell.border_left.color.rgb == %s" % actual


@then("cell.border_left.width == Pt({num_lit})")
def then_cell_border_left_width_eq(context, num_lit):
    actual = context.cell.border_left.width
    expected = Pt(float(num_lit))
    assert actual == expected, "cell.border_left.width == %s EMU" % actual


@then("cell.border_left.dash_style == MSO_LINE.{name}")
def then_cell_border_left_dash_style_eq(context, name):
    actual = context.cell.border_left.dash_style
    expected = getattr(MSO_LINE, name)
    assert actual == expected, "cell.border_left.dash_style == %s" % actual


@then("cell.margin_{side} == Inches({num_lit})")
def then_cell_margin_side_eq_Inches_num(context, side, num_lit):
    actual = getattr(context.cell, "margin_%s" % side)
    expected = Inches(float(num_lit))
    assert actual == expected, "cell.margin_%s == %s" % (side, actual.inches)


@then("{cell_ref}.is_merge_origin is {bool_lit}")
def then_cell_ref_is_merge_origin_is(context, cell_ref, bool_lit):
    expected = eval(bool_lit)
    actual = getattr(context, cell_ref).is_merge_origin
    assert actual is expected, "%s.is_merge_origin is %s" % (cell_ref, actual)


@then("{cell_ref}.is_spanned is {bool_lit}")
def then_cell_is_spanned_is(context, cell_ref, bool_lit):
    expected = eval(bool_lit)
    actual = getattr(context, cell_ref).is_spanned
    assert actual is expected, "%s.is_spanned is %s" % (cell_ref, actual)


@then("cell.text == {value}")
def then_cell_text_eq_value(context, value):
    actual, expected = context.cell.text, eval(value)
    assert actual == expected, 'cell.text == "%s"' % actual


@then("cell.span_height == {int_lit}")
def then_cell_span_height_eq(context, int_lit):
    expected = int(int_lit)
    actual = context.cell.span_height
    assert actual is expected, "cell.span_height == %s" % actual


@then("cell.span_width == {int_lit}")
def then_cell_span_width_eq(context, int_lit):
    expected = int(int_lit)
    actual = context.cell.span_width
    assert actual is expected, "cell.span_width == %s" % actual


@then("origin_cell.span_height == {int_lit}")
def then_origin_cell_span_height_eq(context, int_lit):
    expected = int(int_lit)
    actual = context.origin_cell.span_height
    assert actual == expected, "origin_cell.span_height == %s" % actual


@then("origin_cell.span_width == {int_lit}")
def then_origin_cell_span_width_eq(context, int_lit):
    expected = int(int_lit)
    actual = context.origin_cell.span_width
    assert actual == expected, "origin_cell.span_width == %s" % actual


@then("cell.vertical_anchor == {value}")
def then_cell_vertical_anchor_eq_value(context, value):
    actual = context.cell.vertical_anchor
    expected = eval(value)
    assert actual == expected, "cell.vertical_anchor == %s" % actual


@then("column.width.inches == {float_lit}")
def then_column_width_inches_eq(context, float_lit):
    actual = context.column.width.inches
    expected = float(float_lit)
    assert actual == expected, "column.width.inches == %s" % actual


@then("len(list(table.iter_cells())) == {int_lit}")
def then_len_list_table_iter_cells_eq(context, int_lit):
    actual = len(list(context.table_.iter_cells()))
    expected = int(int_lit)
    assert actual is expected, "len(list(table.iter_cells())) == %s" % actual


@then("len(table.rows) == {int_lit}")
def then_len_table_rows_eq(context, int_lit):
    actual = len(context.table_.rows)
    expected = int(int_lit)
    assert actual == expected, "len(table.rows) == %s" % actual


@then("len(table.columns) == {int_lit}")
def then_len_table_columns_eq(context, int_lit):
    actual = len(context.table_.columns)
    expected = int(int_lit)
    assert actual == expected, "len(table.columns) == %s" % actual


@then("the new row has two cells")
def then_new_row_has_two_cells(context):
    actual = len(list(context.added_row.cells))
    assert actual == 2, "the new row has %d cell(s)" % actual


@then("the new row height equals the prior last row height")
def then_new_row_height_equals_prior(context):
    actual = context.added_row.height
    expected = context.prior_last_row_height
    assert actual == expected, "new row height %s != prior %s" % (actual, expected)


@then("every row has {count} cells")
@then("every row has {count} cell")
def then_every_row_has_count_cells(context, count):
    word_to_int = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5}
    expected = word_to_int.get(count, None)
    if expected is None:
        expected = int(count)
    for idx, row in enumerate(context.table_.rows):
        actual = len(list(row.cells))
        assert actual == expected, "row %d has %d cell(s), expected %d" % (
            idx,
            actual,
            expected,
        )


@then("the new column width equals the prior last column width")
def then_new_column_width_equals_prior(context):
    actual = context.added_column.width
    expected = context.prior_last_column_width
    assert actual == expected, "new column width %s != prior %s" % (actual, expected)


@then("origin_cell.text == {value}")
def then_origin_cell_text_eq_value(context, value):
    actual, expected = context.origin_cell.text, eval(value)
    assert actual == expected, 'origin_cell.text == "%s"' % actual


@then("other_cell.text == {value}")
def then_other_cell_text_eq_value(context, value):
    actual, expected = context.other_cell.text, eval(value)
    assert actual == expected, 'other_cell.text == "%s"' % actual


@then("table.cell(0, 0) is a {type_name} object")
def then_table_cell_0_0_is_a_type_object(context, type_name):
    actual = type(context.table_.cell(0, 0)).__name__
    expected = type_name
    assert actual == expected, "table.cell(0, 0) is a %s object" % actual


@then("table.columns is a {type_name} object")
def then_table_columns_is_a_type_object(context, type_name):
    actual = type(context.table_.columns).__name__
    expected = type_name
    assert actual == expected, "table.columns is a %s object" % actual


@then("table.first_col is {bool_lit}")
def then_table_first_col_is_value(context, bool_lit):
    actual = context.table_.first_col
    expected = eval(bool_lit)
    assert actual is expected, "table.first_col is %s" % actual


@then("table.first_row is {bool_lit}")
def then_table_first_row_is_value(context, bool_lit):
    actual = context.table_.first_row
    expected = eval(bool_lit)
    assert actual is expected, "table.first_row is %s" % actual


@then("table.horz_banding is {bool_lit}")
def then_table_horz_banding_is_value(context, bool_lit):
    actual = context.table_.horz_banding
    expected = eval(bool_lit)
    assert actual is expected, "table.horz_banding is %s" % actual


@then("table.last_col is {bool_lit}")
def then_table_last_col_is_value(context, bool_lit):
    actual = context.table_.last_col
    expected = eval(bool_lit)
    assert actual is expected, "table.last_col is %s" % actual


@then("table.last_row is {bool_lit}")
def then_table_last_row_is_value(context, bool_lit):
    actual = context.table_.last_row
    expected = eval(bool_lit)
    assert actual is expected, "table.last_row is %s" % actual


@then("table.rows is a {type_name} object")
def then_table_rows_is_a_type_object(context, type_name):
    actual = type(context.table_.rows).__name__
    expected = type_name
    assert actual == expected, "table.rows is a %s object" % actual


@then("table.style_id is {value}")
def then_table_style_id_is_value(context, value):
    actual = context.table_.style_id
    expected = eval(value)
    assert actual == expected, "table.style_id is %r" % (actual,)


@then("table.vert_banding is {bool_lit}")
def then_table_vert_banding_is_value(context, bool_lit):
    actual = context.table_.vert_banding
    expected = eval(bool_lit)
    assert actual is expected, "table.vert_banding is %s" % actual


@then("table.cell({r:d}, {c:d}).{attr} == {int_lit:d}")
def then_table_cell_r_c_attr_eq(context, r, c, attr, int_lit):
    actual = getattr(context.table_.cell(r, c), attr)
    assert actual == int_lit, "table.cell(%d, %d).%s == %s" % (r, c, attr, actual)


@then('target_cell.text == "{text_value}"')
def then_target_cell_text_eq_literal(context, text_value):
    actual = context.target_cell.text
    assert actual == text_value, "target_cell.text == %r" % (actual,)


@then("target_cell.fill.fore_color.rgb is RGBColor({r}, {g}, {b})")
def then_target_cell_fill_fore_color_rgb_eq(context, r, g, b):
    actual = context.target_cell.fill.fore_color.rgb
    expected = RGBColor(int(r, 16), int(g, 16), int(b, 16))
    assert actual == expected, "target_cell.fill.fore_color.rgb == %s" % actual


@then("target_cell.{border}.color.rgb is RGBColor({r}, {g}, {b})")
def then_target_cell_border_color_rgb_eq(context, border, r, g, b):
    actual = getattr(context.target_cell, border).color.rgb
    expected = RGBColor(int(r, 16), int(g, 16), int(b, 16))
    assert actual == expected, "target_cell.%s.color.rgb == %s" % (border, actual)


@then("target_cell.{border}.width == Pt({num_lit})")
def then_target_cell_border_width_eq(context, border, num_lit):
    actual = getattr(context.target_cell, border).width
    expected = Pt(float(num_lit))
    assert actual == expected, "target_cell.%s.width == %s EMU" % (border, actual)


@then("target_cell.margin_{side} == Inches({num_lit})")
def then_target_cell_margin_side_eq_Inches(context, side, num_lit):
    actual = getattr(context.target_cell, "margin_%s" % side)
    expected = Inches(float(num_lit))
    assert actual == expected, "target_cell.margin_%s == %s" % (side, actual.inches)


@then("target_cell.vertical_anchor == {value}")
def then_target_cell_vertical_anchor_eq(context, value):
    actual = context.target_cell.vertical_anchor
    expected = eval(value)
    assert actual == expected, "target_cell.vertical_anchor == %s" % actual


@then("target_cell.clone_from returned target_cell")
def then_clone_from_returned_target(context):
    assert context.clone_from_result is context.target_cell, (
        "clone_from did not return target_cell"
    )


@then("source_cell was not mutated by the clone")
def then_source_cell_not_mutated(context):
    actual = context.source_cell._tc.xml  # pyright: ignore[reportPrivateUsage]
    expected = context.source_xml_snapshot
    assert actual == expected, "source_cell XML changed after clone_from"
