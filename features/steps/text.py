# encoding: utf-8

"""
Gherkin step implementations for text-related features.
"""

from __future__ import absolute_import

import os

from behave import given, when, then

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from .helpers import saved_pptx_path, test_pptx


# given ===================================================

@given('a paragraph')
def given_a_paragraph(context):
    context.prs = Presentation()
    blank_slide_layout = context.prs.slide_layouts[6]
    slide = context.prs.slides.add_slide(blank_slide_layout)
    length = Inches(2.00)
    textbox = slide.shapes.add_textbox(length, length, length, length)
    context.paragraph = textbox.textframe.paragraphs[0]


@given('a paragraph containing text')
def given_a_paragraph_containing_text(context):
    prs = Presentation(test_pptx('txt-text'))
    context.paragraph = prs.slides[0].shapes[0].textframe.paragraphs[0]


@given('a run')
def given_a_run(context):
    prs = Presentation(test_pptx('txt-text'))
    context.run = prs.slides[0].shapes[0].textframe.paragraphs[0].runs[0]


@given('a run containing text')
def given_a_run_containing_text(context):
    prs = Presentation(test_pptx('txt-text'))
    context.run = prs.slides[0].shapes[0].textframe.paragraphs[0].runs[0]


@given('a run with italics set {setting}')
def given_run_with_italics_set_to_setting(context, setting):
    run_idx = {'on': 0, 'off': 1, 'to None': 2}[setting]
    context.prs = Presentation(test_pptx('txt-font-props'))
    runs = context.prs.slides[0].shapes[0].textframe.paragraphs[0].runs
    context.run = runs[run_idx]


@given('a text run')
def given_a_text_run(context):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    textbox = slide.shapes.add_textbox(0, 0, 0, 0)
    p = textbox.textframe.paragraphs[0]
    context.r = p.add_run()


@given('a text run in a table cell')
def given_a_text_run_in_a_table_cell(context):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    table = slide.shapes.add_table(1, 1, 0, 0, 0, 0).table
    cell = table.cell(0, 0)
    p = cell.textframe.paragraphs[0]
    context.r = p.add_run()


@given('a text run having a hyperlink')
def given_a_text_run_having_a_hyperlink(context):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    textbox = slide.shapes.add_textbox(0, 0, 0, 0)
    p = textbox.textframe.paragraphs[0]
    r = p.add_run()
    r.hyperlink.address = 'http://foo/bar'
    context.r = r


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
        context.textframe.margin_left = emu
    elif side == 'top':
        context.textframe.margin_top = emu
    elif side == 'right':
        context.textframe.margin_right = emu
    elif side == 'bottom':
        context.textframe.margin_bottom = emu


@when('I indent the paragraph')
def when_indent_first_paragraph(context):
    context.paragraph.level = 1


@when("I set italics {setting}")
def when_set_italics_to_setting(context, setting):
    new_italics_value = {'on': True, 'off': False, 'to None': None}[setting]
    context.run.font.italic = new_italics_value


@when('I set the hyperlink address')
def when_set_hyperlink_address(context):
    context.run_text = 'python-pptx @ GitHub'
    context.address = 'https://github.com/scanny/python-pptx'

    r = context.r
    r.text = context.run_text
    hlink = r.hyperlink
    hlink.address = context.address


@when('I set the hyperlink address to None')
def when_set_hyperlink_address_to_None(context):
    context.r.hyperlink.address = None


@when("I set the paragraph alignment to centered")
def when_set_paragraph_alignment_to_centered(context):
    context.paragraph.alignment = PP_ALIGN.CENTER


# then ====================================================

@then('paragraph.text is the text in the paragraph')
def then_paragraph_text_is_the_text_in_the_paragraph(context):
    paragraph = context.paragraph
    assert paragraph.text == ' Foo Bar \n Baz Zoo \n1'


@then('paragraph.text matches the assigned string')
def then_paragraph_text_matches_the_assigned_string(context):
    paragraph = context.paragraph
    assert paragraph.text == ' Boo Far \n Faz Foo '


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
    p = context.prs.slides[0].shapes[0].textframe.paragraphs[0]
    assert p.alignment == PP_ALIGN.CENTER


@then('the paragraph is indented to the second level')
def then_paragraph_indented_to_second_level(context):
    p = context.prs.slides[0].shapes[0].textframe.paragraphs[0]
    assert p.level == 1


@then("the run that had italics set {initial} now has it set {setting}")
def then_run_now_has_italics_set_to_setting(context, initial, setting):
    run_idx = {'on': 0, 'off': 1, 'to None': 2}[initial]
    run = (
        context.prs
               .slides[0]
               .shapes[0]
               .textframe
               .paragraphs[0]
               .runs[run_idx]
    )
    expected_val = {'on': True, 'off': False, 'to None': None}[setting]
    assert run.font.italic == expected_val


@then('the text of the run is a hyperlink')
def then_text_of_run_is_hyperlink(context):
    r = context.r
    hlink = r.hyperlink
    assert r.text == context.run_text
    assert hlink.address == context.address


@then('the text run is not a hyperlink')
def then_text_run_is_not_a_hyperlink(context):
    hlink = context.r.hyperlink
    assert hlink.address is None
