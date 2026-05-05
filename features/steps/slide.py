"""Gherkin step implementations for slide-related features."""

from __future__ import annotations

import io

from behave import given, then, when
from helpers import test_pptx

from pptx import Presentation

# given ===================================================


@given("a blank slide")
def given_a_blank_slide(context):
    context.prs = Presentation(test_pptx("sld-blank"))
    context.slide = context.prs.slides[0]


@given("a notes slide")
def given_a_notes_slide(context):
    prs = Presentation(test_pptx("sld-notes"))
    context.notes_slide = prs.slides[0].notes_slide


@given("a slide")
def given_a_slide(context):
    presentation = Presentation(test_pptx("shp-shapes"))
    context.slide = presentation.slides[0]


@given("a slide having a notes slide")
def given_a_slide_having_a_notes_slide(context):
    context.slide = Presentation(test_pptx("sld-notes")).slides[0]


@given("a slide having no notes slide")
def given_a_slide_having_no_notes_slide(context):
    context.slide = Presentation(test_pptx("sld-notes")).slides[1]


@given("a slide having a title")
def given_a_slide_having_a_title(context):
    prs = Presentation(test_pptx("shp-shapes"))
    context.prs, context.slide = prs, prs.slides[0]


@given("a slide having name {name}")
def given_a_slide_having_name_name(context, name):
    slide_idx = 0 if name == "Overview" else 1
    presentation = Presentation(test_pptx("sld-slide"))
    context.slide = presentation.slides[slide_idx]


@given("a slide having slide id 256")
def given_a_slide_having_slide_id_256(context):
    presentation = Presentation(test_pptx("shp-shapes"))
    context.slide = presentation.slides[0]


@given("a Slide object based on slide_layout as slide")
def given_a_Slide_object_based_on_slide_layout_as_slide(context):
    context.slide = context.prs.slides[0]


@given("a Slide object having {def_or_ovr} background as slide")
def given_a_Slide_object_having_background_as_slide(context, def_or_ovr):
    slide_idx = {"the default": 0, "an overridden": 1}[def_or_ovr]
    context.slide = Presentation(test_pptx("sld-slide")).slides[slide_idx]


@given("a SlideLayout object as slide")
@given("a SlideLayout object as slide_layout")
def given_a_SlideLayout_object_as_slide(context):
    prs = Presentation(test_pptx("sld-slide"))
    context.slide = context.slide_layout = prs.slide_layouts[0]


@given("a SlideLayout object having name {name} as slide")
def given_a_SlideLayout_object_having_name_as_slide(context, name):
    slide_layout_idx = 0 if name == "of no explicit value" else 1
    prs = Presentation(test_pptx("sld-slide"))
    context.slide = prs.slide_layouts[slide_layout_idx]


@given("a SlideLayout object used by {which_slides} as slide_layout")
def given_a_SlideLayout_object_used_by_slides_as_slide_layout(context, which_slides):
    slide_layout_idx = {"a slide": 0, "no slides": 1}[which_slides]
    context.prs = Presentation(test_pptx("sld-slide"))
    context.slide_layout = context.prs.slide_layouts[slide_layout_idx]


@given("a SlideMaster object as slide")
@given("a SlideMaster object as slide_master")
def given_a_SlideMaster_object_as_slide(context):
    prs = Presentation(test_pptx("sld-slide"))
    context.slide = context.slide_master = prs.slide_masters[0]


@given("a presentation with {n_masters} slide master and no explicit master name")
@given("a presentation with {n_masters} slide masters and no explicit master name")
def given_a_presentation_with_n_masters_and_no_explicit_master_name(context, n_masters):
    fixture = "sld-slide" if int(n_masters) == 1 else "prs-slide-masters"
    context.prs = Presentation(test_pptx(fixture))
    assert len(context.prs.slide_masters) == int(n_masters), (
        "fixture %s has %d masters, expected %s"
        % (fixture, len(context.prs.slide_masters), n_masters)
    )


@given("a fresh presentation with a slide on layout 0")
def given_a_fresh_presentation_with_a_slide_on_layout_0(context):
    context.prs = Presentation()
    context.slide = context.prs.slides.add_slide(context.prs.slide_layouts[0])


@given("a presentation with two masters and a slide on master 0 layout 0")
def given_a_presentation_with_two_masters_and_slide_on_master0_layout0(context):
    context.prs = Presentation(test_pptx("prs-slide-masters"))
    layout = context.prs.slide_masters[0].slide_layouts[0]
    context.slide = context.prs.slides.add_slide(layout)


# when ====================================================


@when('I set prs.slide_masters[{idx:d}].name to "{value}"')
def when_I_set_prs_slide_masters_idx_name_to_value(context, idx, value):
    context.prs.slide_masters[idx].name = value


@when("I set prs.slide_masters[{idx:d}].name to None")
def when_I_set_prs_slide_masters_idx_name_to_None(context, idx):
    context.prs.slide_masters[idx].name = None


@when("I call slide.follow_master_background()")
def when_I_call_slide_follow_master_background(context):
    context.follow_master_bg_result = context.slide.follow_master_background()


@when("I set slide.is_hidden to {value}")
def when_I_set_slide_is_hidden_to_value(context, value):
    # -- ensure we have a Presentation for a later save/load round-trip --
    if not hasattr(context, "prs") or context.prs is None:
        context.prs = Presentation(test_pptx("shp-shapes"))
        context.slide = context.prs.slides[0]
    context.slide.is_hidden = {"True": True, "False": False}[value]


@when("I set slide.show_master_shapes to {value}")
def when_I_set_slide_show_master_shapes_to_value(context, value):
    # -- ensure we have a Presentation for a later save/load round-trip --
    if not hasattr(context, "prs") or context.prs is None:
        context.prs = Presentation(test_pptx("shp-shapes"))
        context.slide = context.prs.slides[0]
    context.slide.show_master_shapes = {"True": True, "False": False}[value]


@when("I set slide.slide_layout to layout {idx:d} of the same master")
def when_I_set_slide_slide_layout_to_layout_N_of_same_master(context, idx):
    context.slide.slide_layout = context.prs.slide_layouts[idx]


@when("I set slide.slide_layout to master {master_idx:d} layout {layout_idx:d}")
def when_I_set_slide_slide_layout_to_master_M_layout_N(context, master_idx, layout_idx):
    layout = context.prs.slide_masters[master_idx].slide_layouts[layout_idx]
    context.slide.slide_layout = layout


# then ====================================================


@then("prs.slide_masters[{idx:d}].name is {expected}")
def then_prs_slide_masters_idx_name_is(context, idx, expected):
    actual = context.prs.slide_masters[idx].name
    assert actual == expected, "expected %r, got %r" % (expected, actual)


@then("after a save/load round-trip prs.slide_masters[{idx:d}].name is {expected}")
def then_after_save_load_prs_slide_masters_idx_name_is(context, idx, expected):
    buf = io.BytesIO()
    context.prs.save(buf)
    buf.seek(0)
    reopened = Presentation(buf)
    actual = reopened.slide_masters[idx].name
    assert actual == expected, "expected %r, got %r" % (expected, actual)


@then("len(notes_slide.shapes) is {count}")
def then_len_notes_slide_shapes_is_count(context, count):
    shapes = context.notes_slide.shapes
    assert len(shapes) == int(count)


@then("notes_slide.notes_placeholder is a NotesSlidePlaceholder object")
def then_notes_slide_notes_placeholder_is_a_NotesSlidePlacehldr_obj(context):
    notes_slide = context.notes_slide
    cls_name = type(notes_slide.notes_placeholder).__name__
    assert cls_name == "NotesSlidePlaceholder", "got %s" % cls_name


@then("notes_slide.notes_text_frame is a TextFrame object")
def then_notes_slide_notes_text_frame_is_a_TextFrame_object(context):
    notes_slide = context.notes_slide
    cls_name = type(notes_slide.notes_text_frame).__name__
    assert cls_name == "TextFrame", "got %s" % cls_name


@then("notes_slide.placeholders is a NotesSlidePlaceholders object")
def then_notes_slide_placeholders_is_a_NotesSlidePlaceholders_object(context):
    notes_slide = context.notes_slide
    assert type(notes_slide.placeholders).__name__ == "NotesSlidePlaceholders"


@then("slide in slide_layout.used_by_slides is True")
def then_slide_in_slide_layout_used_by_slides_is_True(context):
    assert context.slide in context.slide_layout.used_by_slides


@then("slide.background is a _Background object")
def then_slide_background_is_a_Background_object(context):
    cls_name = context.slide.background.__class__.__name__
    assert cls_name == "_Background", "slide.background is a %s object" % cls_name


@then("slide.follow_master_background is {value}")
def then_slide_follow_master_background_is_value(context, value):
    expected_value = {"True": True, "False": False}[value]
    # -- `follow_master_background` is a dual bool-like/callable proxy; cast to bool --
    actual_value = bool(context.slide.follow_master_background)
    assert actual_value is expected_value, "slide.follow_master_background is %s" % actual_value


@then("the call returned the slide itself")
def then_the_call_returned_the_slide_itself(context):
    assert context.follow_master_bg_result is context.slide, (
        "slide.follow_master_background() returned %r, expected the Slide itself"
        % context.follow_master_bg_result
    )


@then("slide.has_notes_slide is {value}")
def then_slide_has_notes_slide_is_value(context, value):
    expected_value = {"True": True, "False": False}[value]
    slide = context.slide
    assert slide.has_notes_slide is expected_value


@then("slide.is_hidden is {value}")
def then_slide_is_hidden_is_value(context, value):
    expected_value = {"True": True, "False": False}[value]
    actual_value = context.slide.is_hidden
    assert actual_value is expected_value, "slide.is_hidden is %s" % actual_value


@then("after a save/load round-trip slide.is_hidden is {value}")
def then_after_save_load_slide_is_hidden_is(context, value):
    expected_value = {"True": True, "False": False}[value]
    buf = io.BytesIO()
    context.prs.save(buf)
    buf.seek(0)
    reopened = Presentation(buf)
    actual_value = reopened.slides[0].is_hidden
    assert actual_value is expected_value, "slide.is_hidden is %s" % actual_value


@then("slide.show_master_shapes is {value}")
def then_slide_show_master_shapes_is_value(context, value):
    expected_value = {"True": True, "False": False}[value]
    actual_value = context.slide.show_master_shapes
    assert actual_value is expected_value, (
        "slide.show_master_shapes is %s" % actual_value
    )


@then("after a save/load round-trip slide.show_master_shapes is {value}")
def then_after_save_load_slide_show_master_shapes_is(context, value):
    expected_value = {"True": True, "False": False}[value]
    buf = io.BytesIO()
    context.prs.save(buf)
    buf.seek(0)
    reopened = Presentation(buf)
    actual_value = reopened.slides[0].show_master_shapes
    assert actual_value is expected_value, (
        "slide.show_master_shapes is %s" % actual_value
    )


@then("slide.slide_layout is layout {idx:d} of master {master_idx:d}")
def then_slide_slide_layout_is_layout_N_of_master_M(context, idx, master_idx):
    expected = context.prs.slide_masters[master_idx].slide_layouts[idx]
    actual = context.slide.slide_layout
    assert actual is expected, "slide.slide_layout is %r, expected %r" % (actual, expected)


@then("after a save/load round-trip slide.slide_layout is layout {idx:d} of master {master_idx:d}")
def then_after_roundtrip_slide_layout_is_layout_N_of_master_M(context, idx, master_idx):
    buf = io.BytesIO()
    context.prs.save(buf)
    buf.seek(0)
    reopened = Presentation(buf)
    expected = reopened.slide_masters[master_idx].slide_layouts[idx]
    actual = reopened.slides[0].slide_layout
    assert actual is expected, "after round-trip slide.slide_layout is %r" % actual


@then("slide.slide_layout is on master {master_idx:d}")
def then_slide_slide_layout_is_on_master_M(context, master_idx):
    expected_master = context.prs.slide_masters[master_idx]
    actual_master = context.slide.slide_layout.slide_master
    assert actual_master is expected_master, (
        "slide.slide_layout.slide_master is %r, expected %r" % (actual_master, expected_master)
    )


@then("after a save/load round-trip slide.slide_layout is on master {master_idx:d}")
def then_after_roundtrip_slide_layout_is_on_master_M(context, master_idx):
    buf = io.BytesIO()
    context.prs.save(buf)
    buf.seek(0)
    reopened = Presentation(buf)
    expected_master = reopened.slide_masters[master_idx]
    actual_master = reopened.slides[0].slide_layout.slide_master
    assert actual_master is expected_master, (
        "after round-trip slide.slide_layout.slide_master is %r" % actual_master
    )


@then("slide.name is {value}")
def then_slide_name_is_value(context, value):
    expected_name = "" if value == "the empty string" else value
    actual_name = context.slide.name
    assert actual_name == expected_name, "slide.name == %s" % actual_name


@then("slide.notes_slide is a NotesSlide object")
def then_slide_notes_slide_is_a_NotesSlide_object(context):
    notes_slide = context.notes_slide = context.slide.notes_slide
    assert type(notes_slide).__name__ == "NotesSlide"


@then("slide.placeholders is a {clsname} object")
def then_slide_placeholders_is_a_clsname_object(context, clsname):
    actual_clsname = context.slide.placeholders.__class__.__name__
    expected_clsname = clsname
    assert actual_clsname == expected_clsname, "slide.placeholders is a %s object" % actual_clsname


@then("slide.shapes is a {clsname} object")
def then_slide_shapes_is_a_clsname_object(context, clsname):
    actual_clsname = context.slide.shapes.__class__.__name__
    expected_clsname = clsname
    assert actual_clsname == expected_clsname, "slide.shapes is a %s object" % actual_clsname


@then("slide.slide_id is 256")
def then_slide_slide_id_is_256(context):
    slide = context.slide
    assert slide.slide_id == 256


@then("slide.slide_layout is the one passed in the call")
def then_slide_slide_layout_is_the_one_passed_in_the_call(context):
    slide = context.prs.slides[3]
    assert slide.slide_layout == context.slide_layout


@then("slide_layout.slide_master is a SlideMaster object")
def then_slide_layout_slide_master_is_a_SlideMaster_object(context):
    slide_layout = context.slide_layout
    assert type(slide_layout.slide_master).__name__ == "SlideMaster"


@then("slide_master.slide_layouts is a SlideLayouts object")
def then_slide_master_slide_layouts_is_a_SlideLayouts_object(context):
    slide_master = context.slide_master
    assert type(slide_master.slide_layouts).__name__ == "SlideLayouts"


@then("slide_layout.used_by_slides == ()")
def then_slide_layout_used_by_slides_eq_empty_tuple(context):
    assert context.slide_layout.used_by_slides == ()


# Slide.clone_shapes_from (CLO-1) ======================================


@given("a fresh presentation with a source slide and a blank target slide")
def given_fresh_prs_with_source_and_blank_target_slide(context):
    from pptx.enum.shapes import MSO_CONNECTOR_TYPE, MSO_SHAPE
    from pptx.util import Inches

    context.prs = Presentation()
    # -- source slide: layout[0] = title slide (has title + subtitle placeholders) --
    context.source_slide = context.prs.slides.add_slide(context.prs.slide_layouts[0])
    title = context.source_slide.shapes.title
    title.text_frame.text = "Source Title"
    # -- add an autoshape, textbox, and connector to the source --
    context.source_slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(3), Inches(2), Inches(1)
    )
    context.source_slide.shapes.add_textbox(Inches(1), Inches(4), Inches(2), Inches(1))
    context.source_slide.shapes.add_connector(
        MSO_CONNECTOR_TYPE.STRAIGHT, Inches(1), Inches(5), Inches(4), Inches(5)
    )
    # -- target slide: blank layout (no placeholder overlap with source) --
    context.target_slide = context.prs.slides.add_slide(context.prs.slide_layouts[6])
    # -- snapshot source state so we can check it wasn't mutated --
    context.src_shape_ids_before = [s.shape_id for s in context.source_slide.shapes]
    context.src_non_placeholder_count = sum(
        1 for s in context.source_slide.shapes if not s.is_placeholder
    )


@when("I call target_slide.clone_shapes_from(source_slide)")
def when_call_clone_shapes_from_default(context):
    context.tgt_tree_count_before = len(context.target_slide.shapes)
    context.target_slide.clone_shapes_from(context.source_slide)


@when("I call target_slide.clone_shapes_from(source_slide, include_placeholders=False)")
def when_call_clone_shapes_from_no_placeholders(context):
    context.tgt_ph_count_before = len(list(context.target_slide.placeholders))
    context.tgt_tree_count_before = len(context.target_slide.shapes)
    context.target_slide.clone_shapes_from(context.source_slide, include_placeholders=False)


@then("target_slide.shapes contains a clone of every source shape")
def then_target_has_clone_of_every_source_shape(context):
    # -- non-placeholder shapes cloned onto target; blank target has none of
    # -- source's placeholder idxs, so placeholder text copying is a no-op.
    added = len(context.target_slide.shapes) - context.tgt_tree_count_before
    expected = context.src_non_placeholder_count
    assert added == expected, f"expected {expected} non-placeholder clones, got {added}"


@then("target_slide shape ids are unique")
def then_target_shape_ids_are_unique(context):
    tgt_ids = [s.shape_id for s in context.target_slide.shapes]
    assert len(tgt_ids) == len(set(tgt_ids)), f"duplicate shape ids on target: {tgt_ids!r}"


@then("source_slide shapes are unchanged")
def then_source_slide_unchanged(context):
    src_ids_after = [s.shape_id for s in context.source_slide.shapes]
    src_ids_before = context.src_shape_ids_before
    assert (
        src_ids_after == src_ids_before
    ), f"source shape ids changed: before={src_ids_before!r} after={src_ids_after!r}"


@then("target_slide has no placeholders cloned from the source")
def then_target_no_placeholder_clones(context):
    # -- placeholder count unchanged on target when include_placeholders=False --
    tgt_ph_count_after = len(list(context.target_slide.placeholders))
    assert tgt_ph_count_after == context.tgt_ph_count_before
    # -- only the non-placeholder source shapes were cloned --
    added = len(context.target_slide.shapes) - context.tgt_tree_count_before
    assert added == context.src_non_placeholder_count
