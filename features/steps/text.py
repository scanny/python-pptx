# encoding: utf-8

"""
Gherkin step implementations for text-related features.
"""

from __future__ import absolute_import

import os

from behave import given, when, then

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches

from helpers import saved_pptx_path, test_pptx


# given ===================================================

@given('a paragraph')
def given_a_paragraph(context):
    prs = Presentation(test_pptx('txt-text'))
    context.prs = prs
    context.paragraph = prs.slides[0].shapes[0].text_frame.paragraphs[0]


@given('a paragraph containing text')
def given_a_paragraph_containing_text(context):
    prs = Presentation(test_pptx('txt-text'))
    context.paragraph = prs.slides[0].shapes[0].text_frame.paragraphs[0]


@given('a paragraph having line spacing of {setting}')
def given_a_paragraph_having_line_spacing_of_setting(context, setting):
    paragraph_idx = {
        'no explicit setting': 0,
        '1.5 lines':           1,
        '20 pt':               2,
    }[setting]
    prs = Presentation(test_pptx('txt-paragraph-spacing'))
    text_frame = prs.slides[2].shapes[0].text_frame
    context.paragraph = text_frame.paragraphs[paragraph_idx]


@given('a paragraph having space {before_after} of {setting}')
def given_a_paragraph_having_space_before_after_of_setting(
        context, before_after, setting):
    slide_idx = {
        'before': 0,
        'after':  1,
    }[before_after]
    paragraph_idx = {
        'no explicit setting': 0,
        '6 pt':                1,
    }[setting]
    prs = Presentation(test_pptx('txt-paragraph-spacing'))
    text_frame = prs.slides[slide_idx].shapes[0].text_frame
    context.paragraph = text_frame.paragraphs[paragraph_idx]


@given('a run')
def given_a_run(context):
    prs = Presentation(test_pptx('txt-text'))
    context.run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]


@given('a run containing text')
def given_a_run_containing_text(context):
    prs = Presentation(test_pptx('txt-text'))
    context.run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]


@given('a text run')
def given_a_text_run(context):
    prs = Presentation(test_pptx('txt-text'))
    context.run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]


@given('a text run in a table cell')
def given_a_text_run_in_a_table_cell(context):
    prs = Presentation(test_pptx('txt-text'))
    cell = prs.slides[1].shapes[0].table.cell(0, 0)
    context.run = cell.text_frame.paragraphs[0].runs[0]


@given('a text run having a hyperlink')
def given_a_text_run_having_a_hyperlink(context):
    prs = Presentation(test_pptx('txt-text'))
    context.run = prs.slides[0].shapes[0].text_frame.paragraphs[1].runs[0]


# when ====================================================

@when('I assign a string to paragraph.text')
def when_I_assign_a_string_to_paragraph_text(context):
    context.paragraph.text = ' Boo Far \n Faz Foo '


@when('I assign a string to run.text')
def when_I_assign_a_string_to_run_text(context):
    context.run.text = ' Boo Far '


@when('I assign a typeface name to the font')
def when_assign_typeface_name_to_font(context):
    context.font.name = 'Verdana'


@when('I assign None to hyperlink.address')
def when_assign_None_to_hyperlink_address(context):
    context.run.hyperlink.address = None


@when('I assign {value_str} to paragraph.line_spacing')
def when_I_assign_value_to_paragraph_line_spacing(context, value_str):
    value = {
        '1.5':    1.5,
        '2.0':    2.0,
        '254000': Emu(254000),
        '304800': Emu(304800),
        'None':   None,
    }[value_str]
    paragraph = context.paragraph
    paragraph.line_spacing = value


@when('I assign {value_str} to paragraph.space_{before_after}')
def when_I_assign_value_to_paragraph_space_before_after(
        context, value_str, before_after):
    value = {
        '76200': 76200,
        '38100': 38100,
        'None':  None,
    }[value_str]
    attr_name = {
        'before': 'space_before',
        'after':  'space_after',
    }[before_after]
    paragraph = context.paragraph
    setattr(paragraph, attr_name, value)


@when('I reload the presentation')
def when_reload_presentation(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.prs.save(saved_pptx_path)
    context.prs = Presentation(saved_pptx_path)


@when('I set the {side} margin to {inches}"')
def when_set_margin_to_value(context, side, inches):
    emu = Inches(float(inches))
    if side == 'left':
        context.text_frame.margin_left = emu
    elif side == 'top':
        context.text_frame.margin_top = emu
    elif side == 'right':
        context.text_frame.margin_right = emu
    elif side == 'bottom':
        context.text_frame.margin_bottom = emu


@when('I indent the paragraph')
def when_indent_first_paragraph(context):
    context.paragraph.level = 1


@when('I set the hyperlink address')
def when_set_hyperlink_address(context):
    context.run_text = 'python-pptx @ GitHub'
    context.address = 'https://github.com/scanny/python-pptx'

    run = context.run
    run.text = context.run_text
    hlink = run.hyperlink
    hlink.address = context.address


@when("I set the paragraph alignment to centered")
def when_set_paragraph_alignment_to_centered(context):
    context.paragraph.alignment = PP_ALIGN.CENTER


# then ====================================================

@then('paragraph.line_spacing is {value_str}')
def then_paragraph_line_spacing_is_value(context, value_str):
    value = {
        'None':   None,
        '1.0':    1.0,
        '1.5':    1.5,
        '2.0':    2.0,
        '254000': 254000,
        '304800': 304800,
    }[value_str]
    paragraph = context.paragraph
    assert paragraph.line_spacing == value


@then('paragraph.line_spacing.pt {result}')
def then_paragraph_line_spacing_pt_result(context, result):
    value, exception = {
        'raises AttributeError': (None, AttributeError),
        'is 20.0':               (20.0, None),
        'is 24.0':               (24.0, None),
    }[result]
    line_spacing = context.paragraph.line_spacing
    if value is not None:
        assert line_spacing.pt == value
    if exception is not None:
        try:
            line_spacing.pt
            raise AssertionError('did not raise')
        except exception:
            pass


@then('paragraph.space_{before_after} is {value_str}')
def then_paragraph_space_before_is_value(context, before_after, value_str):
    attr_name = {
        'before': 'space_before',
        'after':  'space_after',
    }[before_after]
    value = None if value_str == 'None' else int(value_str)
    paragraph = context.paragraph
    assert getattr(paragraph, attr_name) == value


@then('paragraph.space_{before_after}.pt {result}')
def then_paragraph_space_before_pt_is_value(context, before_after, result):
    attr_name = {
        'before': 'space_before',
        'after':  'space_after',
    }[before_after]
    value, exception = {
        'raises AttributeError': (None, AttributeError),
        '== 6.0':                (6.0,  None),
    }[result]
    space_before_after = getattr(context.paragraph, attr_name)
    if value is not None:
        assert space_before_after.pt == value
    if exception is not None:
        try:
            space_before_after.pt
            raise AssertionError('did not raise')
        except exception:
            pass


@then('paragraph.text is the text in the paragraph')
def then_paragraph_text_is_the_text_in_the_paragraph(context):
    paragraph = context.paragraph
    assert paragraph.text == ' Foo Bar \n Baz Zoo \n1'


@then('paragraph.text matches the assigned string')
def then_paragraph_text_matches_the_assigned_string(context):
    paragraph = context.paragraph
    assert paragraph.text == ' Boo Far \n Faz Foo '


@then('run.text is a hyperlink')
def then_run_text_is_a_hyperlink(context):
    run = context.run
    hlink = run.hyperlink
    assert run.text == context.run_text
    assert hlink.address == context.address


@then('run.text is not a hyperlink')
def then_run_text_is_not_a_hyperlink(context):
    hlink = context.run.hyperlink
    assert hlink.address is None


@then('run.text is the text in the run')
def then_run_text_is_the_text_in_the_run(context):
    run = context.run
    assert run.text == ' Foo Bar '


@then('run.text matches the assigned string')
def then_run_text_matches_the_assigned_string(context):
    run = context.run
    assert run.text == ' Boo Far '


@then('the font name matches the typeface I set')
def then_font_name_matches_typeface_I_set(context):
    assert context.font.name == 'Verdana'


@then('the paragraph is aligned centered')
def then_paragraph_is_aligned_centered(context):
    p = context.prs.slides[0].shapes[0].text_frame.paragraphs[0]
    assert p.alignment == PP_ALIGN.CENTER


@then('the paragraph is indented to the second level')
def then_paragraph_indented_to_second_level(context):
    p = context.prs.slides[0].shapes[0].text_frame.paragraphs[0]
    assert p.level == 1
