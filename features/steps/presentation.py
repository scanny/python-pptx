# encoding: utf-8

"""
Gherkin step implementations for presentation-level features.
"""

from __future__ import absolute_import

import os

from behave import given, when, then
from hamcrest import assert_that, is_, is_not, greater_than
from StringIO import StringIO

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from .helpers import basic_pptx_path, ext_rels_pptx_path, saved_pptx_path


# given ===================================================

@given('a clean working directory')
def step_given_clean_working_dir(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)


@given('a presentation with external relationships')
def given_prs_with_ext_rels(context):
    context.prs = Presentation(ext_rels_pptx_path)


@given('an initialized pptx environment')
def step_given_initialized_pptx_env(context):
    pass


@given('I have an empty presentation open')
def step_given_empty_prs(context):
    context.prs = Presentation()


# when ====================================================

@when('I construct a Presentation instance with no path argument')
def step_when_construct_default_prs(context):
    context.prs = Presentation()


@when('I open a basic PowerPoint presentation')
def step_when_open_basic_pptx(context):
    context.prs = Presentation(basic_pptx_path)


@when('I open a presentation contained in a stream')
def step_when_open_presentation_stream(context):
    with open(basic_pptx_path) as f:
        stream = StringIO(f.read())
    context.prs = Presentation(stream)
    stream.close()


@when('I save and reload the presentation')
def when_save_and_reload_prs(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.prs.save(saved_pptx_path)
    context.prs = Presentation(saved_pptx_path)


@when('I save that stream to a file')
def step_when_save_stream_to_a_file(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.stream.seek(0)
    with open(saved_pptx_path, 'wb') as f:
        f.write(context.stream.read())


@when('I save the presentation')
def step_when_save_presentation(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.prs.save(saved_pptx_path)


@when('I save the presentation to a stream')
def step_when_save_presentation_to_stream(context):
    context.stream = StringIO()
    context.prs.save(context.stream)


# then ====================================================

@then('I receive a presentation based on the default template')
def step_then_receive_prs_based_on_def_tmpl(context):
    prs = context.prs
    assert_that(prs, is_not(None))
    slidemasters = prs.slidemasters
    assert_that(slidemasters, is_not(None))
    assert_that(len(slidemasters), is_(1))
    slide_layouts = slidemasters[0].slide_layouts
    assert_that(slide_layouts, is_not(None))
    assert_that(len(slide_layouts), is_(11))


@then('I see the pptx file in the working directory')
def step_then_see_pptx_file_in_working_dir(context):
    assert_that(os.path.isfile(saved_pptx_path))
    minimum = 30000
    actual = os.path.getsize(saved_pptx_path)
    assert_that(actual, is_(greater_than(minimum)))


@then('the external relationships are still there')
def then_ext_rels_are_preserved(context):
    prs = context.prs
    sld = prs.slides[0]
    rel = sld._rels['rId2']
    assert rel.is_external
    assert rel.reltype == RT.HYPERLINK
    assert rel.target_ref == 'https://github.com/scanny/python-pptx'
