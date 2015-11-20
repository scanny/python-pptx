# encoding: utf-8

"""
Gherkin step implementations for presentation-level features.
"""

from __future__ import absolute_import

import os

from behave import given, when, then

from pptx import Presentation
from pptx.compat import BytesIO
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.presentation import _SlideMasters
from pptx.parts.slidemaster import SlideMaster
from pptx.util import Inches

from helpers import saved_pptx_path, test_pptx


# given ===================================================

@given('a clean working directory')
def given_clean_working_dir(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)


@given('a presentation')
def given_a_presentation(context):
    context.presentation = Presentation(test_pptx('prs-properties'))


@given('a presentation having two slide masters')
def given_presentation_having_two_masters(context):
    context.presentation = Presentation(test_pptx('prs-slide-masters'))


@given('a presentation with external relationships')
def given_prs_with_ext_rels(context):
    context.prs = Presentation(test_pptx('ext-rels'))


@given('a slide master collection containing two masters')
def given_slide_master_collection_containing_two_masters(context):
    prs = Presentation(test_pptx('prs-slide-masters'))
    context.slide_masters = prs.slide_masters


@given('an empty presentation')
def given_an_empty_presentation(context):
    context.prs = Presentation()


@given('an initialized pptx environment')
def given_initialized_pptx_env(context):
    pass


# when ====================================================

@when('I change the slide width and height')
def when_change_slide_width_and_height(context):
    presentation = context.presentation
    presentation.slide_width = Inches(4)
    presentation.slide_height = Inches(3)


@when('I construct a Presentation instance with no path argument')
def when_construct_default_prs(context):
    context.prs = Presentation()


@when('I open a basic PowerPoint presentation')
def when_open_basic_pptx(context):
    context.prs = Presentation(test_pptx('test'))


@when('I open a presentation contained in a stream')
def when_open_presentation_stream(context):
    with open(test_pptx('test'), 'rb') as f:
        stream = BytesIO(f.read())
    context.prs = Presentation(stream)
    stream.close()


@when('I save and reload the presentation')
def when_save_and_reload_prs(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.prs.save(saved_pptx_path)
    context.prs = Presentation(saved_pptx_path)


@when('I save that stream to a file')
def when_save_stream_to_a_file(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.stream.seek(0)
    with open(saved_pptx_path, 'wb') as f:
        f.write(context.stream.read())


@when('I save the presentation')
def when_save_presentation(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.prs.save(saved_pptx_path)


@when('I save the presentation to a stream')
def when_save_presentation_to_stream(context):
    context.stream = BytesIO()
    context.prs.save(context.stream)


# then ====================================================

@then('I can access a slide master by index')
def then_can_access_slide_master_by_index(context):
    slide_masters = context.slide_masters
    for idx in range(2):
        slide_master = slide_masters[idx]
        assert isinstance(slide_master, SlideMaster)


@then('I can access the slide master collection of the presentation')
def then_can_access_slide_masters_of_presentation(context):
    presentation = context.presentation
    slide_masters = presentation.slide_masters
    msg = 'Presentation.slide_masters not instance of _SlideMasters'
    assert isinstance(slide_masters, _SlideMasters), msg


@then('I can iterate over the slide masters')
def then_can_iterate_over_the_slide_masters(context):
    slide_masters = context.slide_masters
    actual_count = 0
    for slide_master in slide_masters:
        actual_count += 1
        assert isinstance(slide_master, SlideMaster)
    assert actual_count == 2


@then('I receive a presentation based on the default template')
def then_receive_prs_based_on_def_tmpl(context):
    prs = context.prs
    assert prs is not None
    slide_masters = prs.slide_masters
    assert slide_masters is not None
    assert len(slide_masters) == 1
    slide_layouts = slide_masters[0].slide_layouts
    assert slide_layouts is not None
    assert len(slide_layouts) == 11


@then('its slide height matches its known value')
def then_slide_height_matches_known_value(context):
    presentation = context.presentation
    assert presentation.slide_height == 6858000


@then('its slide width matches its known value')
def then_slide_width_matches_known_value(context):
    presentation = context.presentation
    assert presentation.slide_width == 9144000


@then('I see the pptx file in the working directory')
def then_see_pptx_file_in_working_dir(context):
    assert os.path.isfile(saved_pptx_path)
    minimum = 30000
    actual = os.path.getsize(saved_pptx_path)
    assert actual > minimum


@then('the external relationships are still there')
def then_ext_rels_are_preserved(context):
    prs = context.prs
    sld = prs.slides[0]
    rel = sld._rels['rId2']
    assert rel.is_external
    assert rel.reltype == RT.HYPERLINK
    assert rel.target_ref == 'https://github.com/scanny/python-pptx'


@then('the length of the slide master collection is 2')
def then_len_of_slide_master_collection_is_2(context):
    presentation = context.presentation
    slide_masters = presentation.slide_masters
    assert len(slide_masters) == 2, (
        'expected len(slide_masters) of 2, got %s' % len(slide_masters)
    )


@then('the slide height matches the new value')
def then_slide_height_matches_new_value(context):
    presentation = context.presentation
    assert presentation.slide_height == Inches(3)


@then('the slide width matches the new value')
def then_slide_width_matches_new_value(context):
    presentation = context.presentation
    assert presentation.slide_width == Inches(4)
