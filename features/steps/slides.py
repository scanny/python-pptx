"""Gherkin step implementations for slide collection-related features."""

from __future__ import annotations

from behave import given, then, when
from helpers import test_pptx

from pptx import Presentation

# given ===================================================


@given("a SlideLayouts object containing 2 layouts as slide_layouts")
def given_a_SlideLayouts_object_containing_2_layouts(context):
    prs = Presentation(test_pptx("mst-slide-layouts"))
    context.slide_layouts = prs.slide_master.slide_layouts


@given("a SlideMasters object containing 2 masters")
def given_a_SlideMasters_object_containing_2_masters(context):
    prs = Presentation(test_pptx("prs-slide-masters"))
    context.slide_masters = prs.slide_masters


@given("a Slides object containing 3 slides")
def given_a_Slides_object_containing_3_slides(context):
    prs = Presentation(test_pptx("sld-slides"))
    context.prs = prs
    context.slides = prs.slides


@given('a Presentation with a layout placeholder renamed to "Agenda Title"')
def given_a_presentation_with_layout_placeholder_renamed(context):
    prs = Presentation(test_pptx("prs-add-slide"))
    layout = prs.slide_masters[0].slide_layouts[0]
    # --- rename the title placeholder on the layout to a custom name ---
    for ph in layout.placeholders:
        if ph.placeholder_format.idx == 0:
            ph.name = "Agenda Title"
            break
    context.prs = prs
    context.slide_layout = layout


# when ====================================================


@when("I call slides.add_slide()")
def when_I_call_slides_add_slide(context):
    context.slide_layout = context.prs.slide_masters[0].slide_layouts[0]
    context.slides.add_slide(context.slide_layout)


@when("I call slides.add_slide() with that layout")
def when_I_call_slides_add_slide_with_that_layout(context):
    context.slide = context.prs.slides.add_slide(context.slide_layout)
@when("I call slides.move_slide(slides[0], 2)")
def when_I_call_slides_move_slide(context):
    slides = context.slides
    context.original_slide = slides[0]
    context.original_slide_id = slides[0].slide_id
    slides.move_slide(slides[0], 2)
@when("I call slides.delete(slides[1])")
def when_I_call_slides_delete(context):
    slides = context.slides
    # -- capture identity of the slide that should survive at the ends --
    context.surviving_slide_ids = (slides[0].slide_id, slides[2].slide_id)
    slides.delete(slides[1])


@when("I call slide_layouts.remove(slide_layouts[1])")
def when_I_call_slide_layouts_remove(context):
    slide_layouts = context.slide_layouts
    slide_layouts.remove(slide_layouts[1])


# then ====================================================


@then("iterating produces 3 NotesSlidePlaceholder objects")
def then_iterating_produces_3_NotesSlidePlaceholder_objects(context):
    idx = -1
    for idx, placeholder in enumerate(context.notes_slide.placeholders):
        typename = type(placeholder).__name__
        assert typename == "NotesSlidePlaceholder", "got %s" % typename
    assert idx == 2


@then("iterating slide_layouts produces 2 SlideLayout objects")
def then_iterating_slide_layouts_produces_2_SlideLayout_objects(context):
    slide_layouts = context.slide_layouts
    idx = -1
    for idx, slide_layout in enumerate(slide_layouts):
        assert type(slide_layout).__name__ == "SlideLayout"
    assert idx == 1


@then("iterating slide_masters produces 2 SlideMaster objects")
def then_iterating_slide_masters_produces_2_SlideMaster_objects(context):
    slide_masters = context.slide_masters
    idx = -1
    for idx, slide_master in enumerate(slide_masters):
        assert type(slide_master).__name__ == "SlideMaster"
    assert idx == 1


@then("iterating slides produces 3 Slide objects")
def then_iterating_slides_produces_3_Slide_objects(context):
    slides = context.slides
    idx = -1
    for idx, slide in enumerate(slides):
        assert type(slide).__name__ == "Slide"
    assert idx == 2


@then("len(slides) is {count}")
def then_len_slides_is_count(context, count):
    slides = context.slides
    assert len(slides) == int(count)


@then("len(slide_layouts) is {n}")
def then_len_slide_layouts_is_2(context, n):
    assert len(context.slide_layouts) == int(n)


@then("len(slide_masters) is 2")
def then_len_slide_masters_is_2(context):
    slide_masters = context.slide_masters
    assert len(slide_masters) == 2


@then("slide_layouts[1] is a SlideLayout object")
def then_slide_layouts_1_is_a_SlideLayout_object(context):
    slide_layouts = context.slide_layouts
    assert type(slide_layouts[1]).__name__ == "SlideLayout"


@then("slide_layouts.get_by_name(slide_layouts[1].name) is slide_layouts[1]")
def then_slide_layouts_get_by_name_is_slide_layout(context):
    slide_layouts = context.slide_layouts
    assert slide_layouts.get_by_name(slide_layouts[1].name) is slide_layouts[1]


@then("slide_layouts.index(slide_layouts[1]) == 1")
def then_slide_layouts_index_is_1(context):
    slide_layouts = context.slide_layouts
    assert slide_layouts.index(slide_layouts[1]) == 1


@then("slide_masters[1] is a SlideMaster object")
def then_slide_masters_1_is_a_SlideMaster_object(context):
    slide_masters = context.slide_masters
    assert type(slide_masters[1]).__name__ == "SlideMaster"


@then("slides.get(256) is slides[0]")
def then_slides_get_256_is_slides_0(context):
    slides = context.slides
    assert slides.get(256) is slides[0]


@then("slides.get(666, default=slides[2]) is slides[2]")
def then_slides_get_666_default_slides_2_is_slides_2(context):
    slides = context.slides
    assert slides.get(666, default=slides[2]) is slides[2]


@then("slides[2] is a Slide object")
def then_slides_2_is_a_Slide_object(context):
    slides = context.slides
    assert type(slides[2]).__name__ == "Slide"

# #644 — partname allocation scenario ----------------------


@given("a Presentation with no slides")
def given_a_Presentation_with_no_slides(context):
    prs = Presentation()
    context.prs = prs
    context.slides = prs.slides
    context.slide_layout = prs.slide_masters[0].slide_layouts[0]


@when("I add {count:d} slides")
def when_I_add_count_slides(context, count):
    layout = context.slide_layout
    context.added_slides = [context.slides.add_slide(layout) for _ in range(count)]


@then("each slide partname is unique")
def then_each_slide_partname_is_unique(context):
    partnames = [s.part.partname for s in context.added_slides]
    assert len(set(partnames)) == len(partnames), (
        "duplicate partname(s) detected: %r" % partnames
    )


@then('slide partnames are "{first}" through "{last}"')
def then_slide_partnames_are_first_through_last(context, first, last):
    partnames = [s.part.partname for s in context.added_slides]
    assert partnames[0] == first, "first partname is %r, expected %r" % (
        partnames[0],
        first,
    )
    assert partnames[-1] == last, "last partname is %r, expected %r" % (
        partnames[-1],
        last,
    )


@when("I reference notes_slide on each of them")
def when_I_reference_notes_slide_on_each(context):
    # Touching `.notes_slide` causes a NotesSlidePart to be created for each
    # slide via `Package.next_partname("/ppt/notesSlides/notesSlide%d.xml")`.
    context.notes_slides = [s.notes_slide for s in context.added_slides]


@then("each notes-slide partname is unique")
def then_each_notes_slide_partname_is_unique(context):
    partnames = [ns.part.partname for ns in context.notes_slides]
    assert len(set(partnames)) == len(partnames), (
        "duplicate notes-slide partname(s) detected: %r" % partnames
    )


@then("notes-slide partnames are numbered 1 through {n:d}")
def then_notes_slide_partnames_are_numbered(context, n):
    partnames = sorted(
        ns.part.partname for ns in context.notes_slides
    )
    expected = [
        "/ppt/notesSlides/notesSlide%d.xml" % i for i in range(1, n + 1)
    ]
    # sorting by string puts "10" before "2"; sort by the numeric suffix instead
    partnames.sort(
        key=lambda p: int(
            p.rsplit("notesSlide", 1)[1].rsplit(".xml", 1)[0]
        )
    )
    assert partnames == expected, "got %r, expected %r" % (partnames, expected)


@then('the new slide has a placeholder named "Agenda Title"')
def then_new_slide_has_agenda_title_placeholder(context):
    names = [ph.name for ph in context.slide.placeholders]
    assert "Agenda Title" in names, "got %r" % names

@then("the slide previously at index 0 is now at index 2")
def then_moved_slide_is_at_index_2(context):
    slides = context.slides
    assert slides[2] is context.original_slide, (
        "expected moved slide at index 2, got %r" % slides[2]
    )


@then("its slide_id is unchanged")
def then_moved_slide_id_is_unchanged(context):
    assert (
        context.original_slide.slide_id == context.original_slide_id
    ), "slide_id changed from %d to %d" % (
        context.original_slide_id,
        context.original_slide.slide_id,

@then("the remaining slides are the originals at indices 0 and 2")
def then_remaining_slides_are_originals(context):
    slides = context.slides
    remaining_ids = tuple(s.slide_id for s in slides)
    assert remaining_ids == context.surviving_slide_ids, (
        "expected remaining slide_ids %r, got %r"
        % (context.surviving_slide_ids, remaining_ids)
    )


@then("the presentation round-trips cleanly after delete")
def then_presentation_round_trips_after_delete(context):
    # -- save to an in-memory stream and reopen to validate the package is still
    # -- internally consistent (no dangling refs, all referenced parts present) --
    import io

    from pptx import Presentation

    buf = io.BytesIO()
    context.prs.save(buf)
    buf.seek(0)
    prs2 = Presentation(buf)
    assert len(prs2.slides) == 2, "expected 2 slides after round-trip, got %d" % len(
        prs2.slides
    )
    reopened_ids = tuple(s.slide_id for s in prs2.slides)
    assert reopened_ids == context.surviving_slide_ids, (
        "expected slide_ids %r after round-trip, got %r"
        % (context.surviving_slide_ids, reopened_ids)
    )
