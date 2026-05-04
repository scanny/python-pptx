"""Step implementations for round-trip coverage round 2.

Spans group shapes, sections, transitions, animations, extended-properties,
password-protected save/reload, chartex passthrough, and cross-presentation
operations (add_slide_from_external / merge).
"""

from __future__ import annotations

import io

from behave import given, then, when

from pptx import Presentation
from pptx.enum.animation import MSO_ANIMATION_TRIGGER, MSO_ANIMATION_TYPE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.transition import (
    PP_TRANSITION_SIDE_DIRECTION,
    PP_TRANSITION_SPEED,
    PP_TRANSITION_TYPE,
)
from pptx.exc import EncryptedPackageError
from pptx.util import Inches


def _blank_prs():
    """Return (presentation, blank slide) using the blank layout (6)."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    return prs, slide


def _round_trip(context):
    buf = io.BytesIO()
    context.prs.save(buf)
    buf.seek(0)
    context.prs = Presentation(buf)


# NB: the `when("I round-trip the presentation")` step is already registered
# in steps/round_trip.py (wave 23-C). We rely on that implementation here.


# ==========================================================================
# Group-shape round-trips
# ==========================================================================


@given("a fresh presentation with a group of an oval, a textbox, and a rectangle")
def given_fresh_prs_group_mixed(context):
    prs, slide = _blank_prs()
    oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(1), Inches(1))
    tb = slide.shapes.add_textbox(Inches(3), Inches(1), Inches(2), Inches(1))
    tb.text_frame.text = "label"
    rect = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(6), Inches(1), Inches(1), Inches(1)
    )
    slide.shapes.add_group_shape([oval, tb, rect])
    context.prs = prs


@then("the reloaded group contains 3 child shapes")
def then_reloaded_group_has_3(context):
    group = context.prs.slides[0].shapes[0]
    assert group.shape_type == MSO_SHAPE_TYPE.GROUP, group.shape_type
    children = list(group.shapes)
    assert len(children) == 3, "got %d children" % len(children)


@then('the reloaded group\'s textbox has text "label"')
def then_reloaded_group_textbox_text(context):
    group = context.prs.slides[0].shapes[0]
    texts = [
        s.text_frame.text
        for s in group.shapes
        if getattr(s, "has_text_frame", False) and s.has_text_frame
    ]
    assert "label" in texts, "textbox text 'label' not found, got %r" % texts


@given("a fresh presentation with a two-level nested group")
def given_fresh_prs_nested_group(context):
    prs, slide = _blank_prs()
    a = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(1), Inches(1))
    b = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2), Inches(1), Inches(1), Inches(1))
    inner = slide.shapes.add_group_shape([a, b])
    c = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(3), Inches(3), Inches(1), Inches(1)
    )
    slide.shapes.add_group_shape([inner, c])
    context.prs = prs


@then("the reloaded outer group has 1 group child and 1 non-group child")
def then_reloaded_outer_has_1_group_1_other(context):
    outer = context.prs.slides[0].shapes[0]
    kinds = [s.shape_type for s in outer.shapes]
    n_groups = sum(1 for k in kinds if k == MSO_SHAPE_TYPE.GROUP)
    n_other = len(kinds) - n_groups
    assert n_groups == 1 and n_other == 1, "got groups=%d other=%d" % (n_groups, n_other)


@then("the reloaded inner group has 2 autoshape children")
def then_reloaded_inner_has_2_autoshapes(context):
    outer = context.prs.slides[0].shapes[0]
    inner = next(s for s in outer.shapes if s.shape_type == MSO_SHAPE_TYPE.GROUP)
    children = list(inner.shapes)
    assert len(children) == 2, "got %d inner children" % len(children)


@given("a fresh presentation with a group of three rectangles named A B C")
def given_fresh_prs_group_named_rects(context):
    prs, slide = _blank_prs()
    shapes = []
    for idx, letter in enumerate(("A", "B", "C")):
        s = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(1 + idx),
            Inches(1),
            Inches(0.5),
            Inches(0.5),
        )
        s.name = letter
        shapes.append(s)
    slide.shapes.add_group_shape(shapes)
    context.prs = prs


@then('the reloaded group\'s child shape names are "A" "B" "C"')
def then_reloaded_group_names_ABC(context):
    group = context.prs.slides[0].shapes[0]
    names = [s.name for s in group.shapes]
    assert names == ["A", "B", "C"], names


@given("a fresh presentation with a group of two autoshapes at known coords")
def given_fresh_prs_group_known_coords(context):
    prs, slide = _blank_prs()
    a = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2), Inches(3), Inches(1), Inches(1))
    b = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(4), Inches(5), Inches(1), Inches(1)
    )
    slide.shapes.add_group_shape([a, b])
    context.prs = prs
    context.expected_left = Inches(2)
    context.expected_top = Inches(3)


@then("the reloaded group.left and group.top equal the minimum child offsets")
def then_reloaded_group_extents(context):
    group = context.prs.slides[0].shapes[0]
    assert group.left == context.expected_left, "left=%r" % group.left
    assert group.top == context.expected_top, "top=%r" % group.top


@given("a fresh presentation with a group of two autoshapes")
def given_fresh_prs_group_two_autoshapes(context):
    prs, slide = _blank_prs()
    a = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(1), Inches(1))
    b = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(3), Inches(1), Inches(1), Inches(1)
    )
    slide.shapes.add_group_shape([a, b])
    context.prs = prs


@when("I call ungroup on the reloaded group")
def when_ungroup_reloaded_group(context):
    group = context.prs.slides[0].shapes[0]
    group.ungroup()


@then("the reloaded slide has 2 top-level autoshape children")
def then_reloaded_slide_has_2_autoshapes_top(context):
    shapes = list(context.prs.slides[0].shapes)
    autoshapes = [s for s in shapes if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]
    assert len(autoshapes) == 2, "got %d autoshapes" % len(autoshapes)


@then("the reloaded slide has 0 groups")
def then_reloaded_slide_no_groups(context):
    shapes = list(context.prs.slides[0].shapes)
    groups = [s for s in shapes if s.shape_type == MSO_SHAPE_TYPE.GROUP]
    assert len(groups) == 0, "got %d groups" % len(groups)


# ==========================================================================
# Sections round-trips
# ==========================================================================


@given('a fresh presentation with a section renamed "Main"')
def given_fresh_prs_section_renamed_main(context):
    prs = Presentation()
    for _ in range(3):
        prs.slides.add_slide(prs.slide_layouts[6])
    prs.sections.add_section("Working", slides=[prs.slides[0]])
    prs.sections[0].name = "Main"
    context.prs = prs


@then('the reloaded presentation has 1 section named "Main"')
def then_reloaded_has_one_section_main(context):
    assert len(context.prs.sections) == 1
    assert context.prs.sections[0].name == "Main"


@when("I remove the only section from the reloaded presentation")
def when_remove_only_section(context):
    only = context.prs.sections[0]
    context.prs.sections.remove(only)


@then("the reloaded presentation has 0 sections")
def then_reloaded_has_zero_sections(context):
    assert len(context.prs.sections) == 0


@given('a fresh presentation with three sections "A" "B" "C"')
def given_fresh_prs_three_sections(context):
    prs = Presentation()
    for _ in range(3):
        prs.slides.add_slide(prs.slide_layouts[6])
    prs.sections.add_section("A")
    prs.sections.add_section("B")
    prs.sections.add_section("C")
    context.prs = prs


@when('I reorder the sections to "C" "A" "B"')
def when_reorder_sections_CAB(context):
    a = context.prs.sections.get_by_name("A")
    b = context.prs.sections.get_by_name("B")
    c = context.prs.sections.get_by_name("C")
    assert a and b and c
    c.move_before(a)
    # -- initial was A B C; after move_before(a) -> C A B --


@then('the reloaded section names are "C", "A", "B"')
def then_reloaded_section_names_CAB(context):
    names = [s.name for s in context.prs.sections]
    assert names == ["C", "A", "B"], names


@given("a fresh presentation with a section authored with a fixed GUID")
def given_fresh_prs_section_fixed_guid(context):
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    context.fixed_guid = "{12345678-1234-1234-1234-123456789ABC}"
    prs.sections.add_section("Fixed", id=context.fixed_guid)
    context.prs = prs


@then("the reloaded first section id equals the fixed GUID")
def then_reloaded_first_section_id_equals(context):
    actual = context.prs.sections[0].id
    assert actual.lower() == context.fixed_guid.lower(), actual


@given("a fresh presentation with two sections splitting four slides 2/2")
def given_fresh_prs_split_2_2(context):
    prs = Presentation()
    for _ in range(4):
        prs.slides.add_slide(prs.slide_layouts[6])
    prs.sections.add_section("First", slides=[prs.slides[0], prs.slides[1]])
    prs.sections.add_section("Second", slides=[prs.slides[2], prs.slides[3]])
    context.prs = prs


@when("I move slide 3 into the first section")
def when_move_slide_3_to_first(context):
    first = context.prs.sections[0]
    first.move_slide(context.prs.slides[2])


@then("the reloaded first section has 3 slides")
def then_reloaded_first_section_three(context):
    slides = context.prs.sections[0].slides
    assert len(slides) == 3, "got %d" % len(slides)


@then("the reloaded second section has 1 slide")
def then_reloaded_second_section_one(context):
    slides = context.prs.sections[1].slides
    assert len(slides) == 1, "got %d" % len(slides)


@given("a fresh presentation with sections that leave slide 3 unassigned")
def given_fresh_prs_unassigned_slide(context):
    prs = Presentation()
    for _ in range(3):
        prs.slides.add_slide(prs.slide_layouts[6])
    prs.sections.add_section("Intro", slides=[prs.slides[0], prs.slides[1]])
    # -- slide 3 intentionally left out; also author no section for it. --
    context.prs = prs


@then("find_containing for slide 3 is None on the reloaded presentation")
def then_find_containing_slide3_none(context):
    section = context.prs.sections.find_containing(context.prs.slides[2])
    assert section is None, "expected None, got %r" % section


# ==========================================================================
# Transitions round-trips
# ==========================================================================


@given("a fresh presentation with transition speed SLOW on slide 0")
def given_fresh_prs_transition_speed_slow(context):
    prs, slide = _blank_prs()
    slide.transition.type = PP_TRANSITION_TYPE.FADE
    slide.transition.speed = PP_TRANSITION_SPEED.SLOW
    context.prs = prs


@then("the reloaded transition.speed is PP_TRANSITION_SPEED.SLOW")
def then_reloaded_transition_speed_slow(context):
    speed = context.prs.slides[0].transition.speed
    assert speed is PP_TRANSITION_SPEED.SLOW, speed


@given("a fresh presentation with a wipe transition direction RIGHT on slide 0")
def given_fresh_prs_wipe_right(context):
    prs, slide = _blank_prs()
    slide.transition.wipe_direction = PP_TRANSITION_SIDE_DIRECTION.RIGHT
    context.prs = prs


@then("the reloaded transition.wipe_direction is PP_TRANSITION_SIDE_DIRECTION.RIGHT")
def then_reloaded_wipe_right(context):
    actual = context.prs.slides[0].transition.wipe_direction
    assert actual is PP_TRANSITION_SIDE_DIRECTION.RIGHT, actual


@given("a fresh presentation with transition advance_after_time set to 5000")
def given_fresh_prs_transition_adv_5000(context):
    prs, slide = _blank_prs()
    slide.transition.advance_after_time = 5000
    context.prs = prs


@when("I clear advance_after_time on that slide")
def when_clear_adv_after_time(context):
    context.prs.slides[0].transition.advance_after_time = None


@then("the reloaded transition.advance_after_time is None")
def then_reloaded_adv_after_time_none(context):
    assert context.prs.slides[0].transition.advance_after_time is None


@given("a fresh presentation with slide 0 FADE and slide 1 MORPH")
def given_fresh_prs_slide0_fade_slide1_morph(context):
    prs = Presentation()
    s0 = prs.slides.add_slide(prs.slide_layouts[6])
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    s0.transition.type = PP_TRANSITION_TYPE.FADE
    s1.transition.type = PP_TRANSITION_TYPE.MORPH
    context.prs = prs


@then("the reloaded slide 0 transition.type is PP_TRANSITION_TYPE.FADE")
def then_reloaded_slide0_fade(context):
    t = context.prs.slides[0].transition.type
    assert t is PP_TRANSITION_TYPE.FADE, t


@then("the reloaded slide 1 transition.type is PP_TRANSITION_TYPE.MORPH")
def then_reloaded_slide1_morph(context):
    t = context.prs.slides[1].transition.type
    assert t is PP_TRANSITION_TYPE.MORPH, t


@given("a fresh presentation with a FADE transition then cleared")
def given_fresh_prs_fade_then_none(context):
    prs, slide = _blank_prs()
    slide.transition.type = PP_TRANSITION_TYPE.FADE
    slide.transition.type = PP_TRANSITION_TYPE.NONE
    context.prs = prs


@then("the reloaded transition.type is PP_TRANSITION_TYPE.NONE")
def then_reloaded_transition_none(context):
    t = context.prs.slides[0].transition.type
    assert t is PP_TRANSITION_TYPE.NONE, t


# ==========================================================================
# Animation round-trips
# ==========================================================================


def _add_rect_with_anim(prs, slide, anim_type, delay=0):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
    )
    shape.set_animation(anim_type, trigger=MSO_ANIMATION_TRIGGER.ON_CLICK, delay=delay)
    return shape


@given("a fresh presentation with a shape carrying a FADE_IN animation")
def given_fresh_prs_fade_in(context):
    prs, slide = _blank_prs()
    _add_rect_with_anim(prs, slide, MSO_ANIMATION_TYPE.FADE_IN, delay=0)
    context.prs = prs


@then("the reloaded shape.animation.type is MSO_ANIMATION_TYPE.FADE_IN")
def then_reloaded_anim_fade_in(context):
    anim = context.prs.slides[0].shapes[0].animation
    assert anim is not None, "anim is None"
    assert anim.type is MSO_ANIMATION_TYPE.FADE_IN, anim.type


@then("the reloaded shape.animation.delay is 0")
def then_reloaded_anim_delay_0(context):
    anim = context.prs.slides[0].shapes[0].animation
    assert anim is not None and anim.delay == 0, anim.delay if anim else None


@given("a fresh presentation with a shape carrying a PULSE emphasis animation")
def given_fresh_prs_pulse(context):
    prs, slide = _blank_prs()
    _add_rect_with_anim(prs, slide, MSO_ANIMATION_TYPE.PULSE)
    context.prs = prs


@then("the reloaded shape.animation.type is MSO_ANIMATION_TYPE.PULSE")
def then_reloaded_anim_pulse(context):
    anim = context.prs.slides[0].shapes[0].animation
    assert anim is not None and anim.type is MSO_ANIMATION_TYPE.PULSE, (
        anim.type if anim else None
    )


@given("a fresh presentation with a shape carrying a FADE_OUT exit animation")
def given_fresh_prs_fade_out(context):
    prs, slide = _blank_prs()
    _add_rect_with_anim(prs, slide, MSO_ANIMATION_TYPE.FADE_OUT)
    context.prs = prs


@then("the reloaded shape.animation.type is MSO_ANIMATION_TYPE.FADE_OUT")
def then_reloaded_anim_fade_out(context):
    anim = context.prs.slides[0].shapes[0].animation
    assert anim is not None and anim.type is MSO_ANIMATION_TYPE.FADE_OUT, (
        anim.type if anim else None
    )


@when("I clear the shape's animation")
def when_clear_shape_anim(context):
    context.prs.slides[0].shapes[0].set_animation(None)


@then("the reloaded shape.animation is None")
def then_reloaded_anim_is_none(context):
    assert context.prs.slides[0].shapes[0].animation is None


@then("the reloaded slide.has_animations is False")
def then_reloaded_has_animations_false(context):
    assert context.prs.slides[0].has_animations is False


@given("a fresh presentation with two shapes authored FADE_IN then PULSE")
def given_fresh_prs_two_anims(context):
    prs, slide = _blank_prs()
    _add_rect_with_anim(prs, slide, MSO_ANIMATION_TYPE.FADE_IN)
    rect2 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(4), Inches(1), Inches(2), Inches(1)
    )
    rect2.set_animation(MSO_ANIMATION_TYPE.PULSE)
    context.prs = prs


@then('the reloaded animation_sequence preset_classes begin with "entr" then "emph"')
def then_reloaded_anim_preset_classes(context):
    seq = context.prs.slides[0].animation_sequence
    classes = [e.preset_class for e in seq]
    # -- dedupe consecutive duplicates (PowerPoint authors each preset as a
    # -- pair of effect entries; we only care about the top-level order) --
    deduped = []
    for c in classes:
        if not deduped or deduped[-1] != c:
            deduped.append(c)
    assert deduped[:2] == ["entr", "emph"], deduped


# ==========================================================================
# Extended-properties round-trips
# ==========================================================================


@given('a fresh presentation with extended company "Acme Corp"')
def given_fresh_prs_ext_company(context):
    prs = Presentation()
    prs.extended_properties.company = "Acme Corp"
    context.prs = prs


@then('the reloaded extended_properties.company is "Acme Corp"')
def then_reloaded_ext_company(context):
    assert context.prs.extended_properties.company == "Acme Corp"


@given('a fresh presentation with extended manager "Jane Doe"')
def given_fresh_prs_ext_manager(context):
    prs = Presentation()
    prs.extended_properties.manager = "Jane Doe"
    context.prs = prs


@then('the reloaded extended_properties.manager is "Jane Doe"')
def then_reloaded_ext_manager(context):
    assert context.prs.extended_properties.manager == "Jane Doe"


@given("a fresh presentation with authored app / app_version / hyperlink_base")
def given_fresh_prs_ext_app(context):
    prs = Presentation()
    prs.extended_properties.application = "python-pptx"
    prs.extended_properties.app_version = "1.0"
    prs.extended_properties.hyperlink_base = "https://example.com/"
    context.prs = prs


@then('the reloaded extended_properties.application is "python-pptx"')
def then_reloaded_ext_application(context):
    assert context.prs.extended_properties.application == "python-pptx"


@then('the reloaded extended_properties.app_version is "1.0"')
def then_reloaded_ext_app_version(context):
    assert context.prs.extended_properties.app_version == "1.0"


@then('the reloaded extended_properties.hyperlink_base is "https://example.com/"')
def then_reloaded_ext_hyperlink_base(context):
    assert context.prs.extended_properties.hyperlink_base == "https://example.com/"


@given("a fresh presentation with two additional blank slides")
def given_fresh_prs_three_blank(context):
    prs = Presentation()
    for _ in range(3):
        prs.slides.add_slide(prs.slide_layouts[6])
    context.prs = prs


@then("the reloaded extended_properties.slide_count equals 3")
def then_reloaded_ext_slide_count(context):
    assert context.prs.extended_properties.slide_count == 3


# ==========================================================================
# Password-protected save/reload
# ==========================================================================


@given('a fresh presentation with a textbox "secret"')
def given_fresh_prs_textbox_secret(context):
    prs, slide = _blank_prs()
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    tb.text_frame.text = "secret"
    context.prs = prs


@when('I save the presentation encrypted with password "{password}"')
def when_save_encrypted(context, password):
    buf = io.BytesIO()
    context.prs.save(buf, password=password)
    context.encrypted_bytes = buf.getvalue()
    context.encryption_password = password


@when("I reopen the encrypted package with the correct password")
def when_reopen_encrypted_correct(context):
    buf = io.BytesIO(context.encrypted_bytes)
    context.prs = Presentation(buf, password=context.encryption_password)


@when('I reopen the encrypted package with password "{password}"')
def when_reopen_encrypted_password(context, password):
    buf = io.BytesIO(context.encrypted_bytes)
    context.prs = Presentation(buf, password=password)


@then('the reloaded textbox text is "secret"')
def then_reloaded_textbox_secret(context):
    tb = context.prs.slides[0].shapes[0]
    assert tb.text_frame.text == "secret", repr(tb.text_frame.text)


@then("opening the package without a password raises EncryptedPackageError")
def then_open_no_password_raises(context):
    buf = io.BytesIO(context.encrypted_bytes)
    try:
        Presentation(buf)
    except EncryptedPackageError:
        return
    raise AssertionError("EncryptedPackageError not raised")


@then("opening the package with the wrong password raises EncryptedPackageError")
def then_open_wrong_password_raises(context):
    buf = io.BytesIO(context.encrypted_bytes)
    try:
        Presentation(buf, password="wrong")
    except EncryptedPackageError:
        return
    raise AssertionError("EncryptedPackageError not raised")


@given('a fresh presentation whose core title is "Encrypted Deck"')
def given_fresh_prs_core_title_encrypted_deck(context):
    prs = Presentation()
    prs.core_properties.title = "Encrypted Deck"
    context.prs = prs


@then('the reloaded presentation core_properties.title is "Encrypted Deck"')
def then_reloaded_core_title(context):
    assert context.prs.core_properties.title == "Encrypted Deck"


# ==========================================================================
# Chartex deep round-trip
# ==========================================================================


@when("I save and reload the chartex presentation twice")
def when_save_reload_chartex_twice(context):
    # -- first cycle --
    buf = io.BytesIO()
    context.prs.save(buf)
    buf.seek(0)
    context.prs = Presentation(buf)
    # -- second cycle --
    buf = io.BytesIO()
    context.prs.save(buf)
    buf.seek(0)
    context.prs = Presentation(buf)


@then("the chartex shapes still number 2 after the second reload")
def then_chartex_shapes_two_after_two_reloads(context):
    shapes = list(context.prs.slides[0].shapes)
    chartex = [s for s in shapes if getattr(s, "has_chartex", False)]
    assert len(chartex) == 2, "got %d" % len(chartex)


@then("every chartex shape on the reloaded slide has shape_type CHART")
def then_every_chartex_is_chart_shape_type(context):
    shapes = list(context.prs.slides[0].shapes)
    chartex = [s for s in shapes if getattr(s, "has_chartex", False)]
    assert len(chartex) >= 1
    for s in chartex:
        assert s.shape_type is MSO_SHAPE_TYPE.CHART, s.shape_type


# ==========================================================================
# Cross-presentation / merge
# ==========================================================================


@given('a source presentation with a slide that has a textbox "hello"')
def given_source_prs_textbox_hello(context):
    src = Presentation()
    slide = src.slides.add_slide(src.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    tb.text_frame.text = "hello"
    context.source = src


@given("a target presentation with the default layout")
def given_target_prs_default_layout(context):
    context.prs = Presentation()


@when("I call add_slide_from_external on the target from the source slide")
def when_add_slide_from_external(context):
    source_slide = context.source.slides[0]
    layout = context.prs.slide_layouts[6]
    context.prs.slides.add_slide_from_external(source_slide, layout)


@when("I round-trip the target presentation")
def when_round_trip_target(context):
    _round_trip(context)


@then('the reloaded target has a slide with a textbox "hello"')
def then_reloaded_target_hello(context):
    slides = list(context.prs.slides)
    found = False
    for slide in slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                if shape.text_frame.text == "hello":
                    found = True
                    break
        if found:
            break
    assert found, "textbox 'hello' not found on reloaded target"


@given("a source presentation with two distinct slides A and B")
def given_source_prs_two_slides_AB(context):
    src = Presentation()
    for text in ("A", "B"):
        slide = src.slides.add_slide(src.slide_layouts[6])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
        tb.text_frame.text = text
    context.source = src


@given("a fresh target presentation")
def given_fresh_target_prs(context):
    context.prs = Presentation()


@when("I call prs.merge(other) from target onto source")
def when_prs_merge_other(context):
    context.prs.merge(context.source)


@then("the reloaded target presentation has 2 slides")
def then_reloaded_target_has_2_slides(context):
    assert len(context.prs.slides) == 2, "got %d" % len(context.prs.slides)


@then('the reloaded target\'s first slide has a textbox "A"')
def then_reloaded_target_first_textbox_A(context):
    shapes = context.prs.slides[0].shapes
    texts = [
        s.text_frame.text
        for s in shapes
        if getattr(s, "has_text_frame", False) and s.has_text_frame
    ]
    assert "A" in texts, texts


@then('the reloaded target\'s second slide has a textbox "B"')
def then_reloaded_target_second_textbox_B(context):
    shapes = context.prs.slides[1].shapes
    texts = [
        s.text_frame.text
        for s in shapes
        if getattr(s, "has_text_frame", False) and s.has_text_frame
    ]
    assert "B" in texts, texts


@given("a fresh presentation with a single blank slide")
def given_fresh_prs_single_blank_slide(context):
    prs, _ = _blank_prs()
    context.prs = prs


@then("calling presentation.merge with itself raises ValueError")
def then_merge_self_raises(context):
    try:
        context.prs.merge(context.prs)
    except ValueError:
        return
    raise AssertionError("ValueError not raised by merge(self)")


@given('a source presentation with a slide whose shape is named "Hero"')
def given_source_prs_shape_named_hero(context):
    src = Presentation()
    slide = src.slides.add_slide(src.slide_layouts[6])
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
    )
    shape.name = "Hero"
    context.source = src


@then('the reloaded target\'s cloned slide has a shape named "Hero"')
def then_reloaded_target_shape_named_hero(context):
    slides = list(context.prs.slides)
    found = False
    for slide in slides:
        for shape in slide.shapes:
            if shape.name == "Hero":
                found = True
                break
        if found:
            break
    assert found, "shape named 'Hero' not found on reloaded target"


