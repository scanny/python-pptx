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
from pptx.util import Inches

from helpers import saved_pptx_path, test_pptx


# given ===================================================


@given("a clean working directory")
def given_clean_working_dir(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)


@given("a presentation")
def given_a_presentation(context):
    context.presentation = Presentation(test_pptx("prs-properties"))


@given("a presentation having a notes master")
def given_a_presentation_having_a_notes_master(context):
    context.prs = Presentation(test_pptx("prs-notes"))


@given("a presentation having no notes master")
def given_a_presentation_having_no_notes_master(context):
    context.prs = Presentation(test_pptx("prs-properties"))


@given("a presentation with external relationships")
def given_prs_with_ext_rels(context):
    context.prs = Presentation(test_pptx("ext-rels"))


@given("an initialized pptx environment")
def given_initialized_pptx_env(context):
    pass


# when ====================================================


@when("I change the slide width and height")
def when_change_slide_width_and_height(context):
    presentation = context.presentation
    presentation.slide_width = Inches(4)
    presentation.slide_height = Inches(3)


@when("I construct a Presentation instance with no path argument")
def when_construct_default_prs(context):
    context.prs = Presentation()


@when("I open a basic PowerPoint presentation")
def when_open_basic_pptx(context):
    context.prs = Presentation(test_pptx("test"))


@when("I open a presentation contained in a stream")
def when_open_presentation_stream(context):
    with open(test_pptx("test"), "rb") as f:
        stream = BytesIO(f.read())
    context.prs = Presentation(stream)
    stream.close()


@when("I save and reload the presentation")
def when_save_and_reload_prs(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.prs.save(saved_pptx_path)
    context.prs = Presentation(saved_pptx_path)


@when("I save that stream to a file")
def when_save_stream_to_a_file(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.stream.seek(0)
    with open(saved_pptx_path, "wb") as f:
        f.write(context.stream.read())


@when("I save the presentation")
def when_save_presentation(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.prs.save(saved_pptx_path)


@when("I save the presentation to a stream")
def when_save_presentation_to_stream(context):
    context.stream = BytesIO()
    context.prs.save(context.stream)


# then ====================================================


@then("I receive a presentation based on the default template")
def then_receive_prs_based_on_def_tmpl(context):
    prs = context.prs
    assert prs is not None
    slide_masters = prs.slide_masters
    assert slide_masters is not None
    assert len(slide_masters) == 1
    slide_layouts = slide_masters[0].slide_layouts
    assert slide_layouts is not None
    assert len(slide_layouts) == 11


@then("its slide height matches its known value")
def then_slide_height_matches_known_value(context):
    presentation = context.presentation
    assert presentation.slide_height == 6858000


@then("its slide width matches its known value")
def then_slide_width_matches_known_value(context):
    presentation = context.presentation
    assert presentation.slide_width == 9144000


@then("I see the pptx file in the working directory")
def then_see_pptx_file_in_working_dir(context):
    assert os.path.isfile(saved_pptx_path)
    minimum = 30000
    actual = os.path.getsize(saved_pptx_path)
    assert actual > minimum


@then("len(notes_master.shapes) is {shape_count}")
def then_len_notes_master_shapes_is_shape_count(context, shape_count):
    notes_master = context.prs.notes_master
    expected = int(shape_count)
    actual = len(notes_master.shapes)
    assert actual == expected, "got %s" % actual


@then("prs.notes_master is a NotesMaster object")
def then_prs_notes_master_is_a_NotesMaster_object(context):
    prs = context.prs
    assert type(prs.notes_master).__name__ == "NotesMaster"


@then("prs.slides is a Slides object")
def then_prs_slides_is_a_Slides_object(context):
    prs = context.presentation
    assert type(prs.slides).__name__ == "Slides"


@then("prs.slide_masters is a SlideMasters object")
def then_prs_slide_masters_is_a_SlideMasters_object(context):
    prs = context.presentation
    assert type(prs.slide_masters).__name__ == "SlideMasters"


@then("the external relationships are still there")
def then_ext_rels_are_preserved(context):
    prs = context.prs
    sld = prs.slides[0]
    rel = sld.part._rels["rId2"]
    assert rel.is_external
    assert rel.reltype == RT.HYPERLINK
    assert rel.target_ref == "https://github.com/scanny/python-pptx"


@then("the slide height matches the new value")
def then_slide_height_matches_new_value(context):
    presentation = context.presentation
    assert presentation.slide_height == Inches(3)


@then("the slide width matches the new value")
def then_slide_width_matches_new_value(context):
    presentation = context.presentation
    assert presentation.slide_width == Inches(4)
