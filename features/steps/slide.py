# encoding: utf-8

"""
Gherkin step implementations for slide-related features.
"""

from __future__ import absolute_import

from behave import given, when, then

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.shapes.base import BaseShape

from helpers import test_pptx

SHAPE_COUNT = 3


# given ===================================================

@given('a blank slide')
def given_a_blank_slide(context):
    context.prs = Presentation(test_pptx('sld-blank'))
    context.slide = context.prs.slides[0]


@given('a LayoutPlaceholders object containing 2 placeholders')
def given_a_LayoutPlaceholders_object_containing_2_placeholders(context):
    prs = Presentation(test_pptx('lyt-shapes'))
    context.layout_placeholders = prs.slide_layouts[0].placeholders


@given('a LayoutShapes object containing 3 shapes')
def given_a_LayoutShapes_object_containing_3_shapes(context):
    prs = Presentation(test_pptx('lyt-shapes'))
    context.layout_shapes = prs.slide_layouts[0].shapes


@given('a MasterPlaceholders object containing 2 placeholders')
def given_a_MasterPlaceholders_object_containing_2_placeholders(context):
    prs = Presentation(test_pptx('mst-placeholders'))
    context.master_placeholders = prs.slide_masters[0].placeholders


@given('a MasterShapes object containing 2 shapes')
def given_a_MasterShapes_object_containing_2_shapes(context):
    prs = Presentation(test_pptx('mst-shapes'))
    context.master_shapes = prs.slide_masters[0].shapes


@given('a notes slide')
def given_a_notes_slide(context):
    prs = Presentation(test_pptx('sld-notes'))
    context.notes_slide = prs.slides[0].notes_slide


@given('a slide')
def given_a_slide(context):
    presentation = Presentation(test_pptx('shp-shape-access'))
    context.slide = presentation.slides[0]


@given('a slide having a notes slide')
def given_a_slide_having_a_notes_slide(context):
    context.slide = Presentation(test_pptx('sld-notes')).slides[0]


@given('a slide having no notes slide')
def given_a_slide_having_no_notes_slide(context):
    context.slide = Presentation(test_pptx('sld-notes')).slides[1]


@given('a slide having a title')
def given_a_slide_having_a_title(context):
    prs = Presentation(test_pptx('shp-shape-access'))
    context.prs, context.slide = prs, prs.slides[0]


@given('a slide having name {name}')
def given_a_slide_having_name_name(context, name):
    slide_idx = 0 if name == 'Overview' else 1
    presentation = Presentation(test_pptx('sld-slide-props'))
    context.slide = presentation.slides[slide_idx]


@given('a slide having slide id 256')
def given_a_slide_having_slide_id_256(context):
    presentation = Presentation(test_pptx('shp-shape-access'))
    context.slide = presentation.slides[0]


@given('a slide layout')
def given_a_slide_layout(context):
    prs = Presentation(test_pptx('sld-slide-access'))
    context.slide_layout = prs.slide_layouts[0]


@given('a slide layout having name {name}')
def given_a_slide_layout_having_name_name(context, name):
    slide_layout_idx = 0 if name == 'of no explicit value' else 1
    presentation = Presentation(test_pptx('sld-slide-props'))
    context.slide_layout = presentation.slide_layouts[slide_layout_idx]


@given('a slide master')
def given_a_slide_master(context):
    prs = Presentation(test_pptx('sld-slide-access'))
    context.slide_master = prs.slide_masters[0]


@given('a SlideLayouts object containing 2 layouts')
def given_a_SlideLayouts_object_containing_2_layouts(context):
    prs = Presentation(test_pptx('mst-slide-layouts'))
    context.slide_layouts = prs.slide_master.slide_layouts


@given('a SlideMasters object containing 2 masters')
def given_a_SlideMasters_object_containing_2_masters(context):
    prs = Presentation(test_pptx('prs-slide-masters'))
    context.slide_masters = prs.slide_masters


@given('a SlidePlaceholders object containing 2 placeholders')
def given_a_SlidePlaceholders_object_containing_2_placeholders(context):
    prs = Presentation(test_pptx('shp-shape-access'))
    context.placeholders = prs.slides[0].placeholders


@given('a Slides object containing 3 slides')
def given_a_Slides_object_containing_3_slides(context):
    prs = Presentation(test_pptx('sld-slide-access'))
    context.prs = prs
    context.slides = prs.slides


# when ====================================================

@when('I call slides.add_slide()')
def when_I_call_slides_add_slide(context):
    context.slide_layout = context.prs.slide_masters[0].slide_layouts[0]
    context.slides.add_slide(context.slide_layout)


# then ====================================================

@then('iterating layout_placeholders produces 2 LayoutPlaceholder objects')
def then_iterating_layout_placeholders_produces_2_objects(context):
    layout_placeholders = context.layout_placeholders
    idx = -1
    for idx, layout_placeholder in enumerate(layout_placeholders):
        assert type(layout_placeholder).__name__ == 'LayoutPlaceholder'
    assert idx == 1


@then('iterating layout_shapes produces 3 BaseShape objects')
def then_iterating_layout_shapes_produces_3_BaseShape_objects(context):
    layout_shapes = context.layout_shapes
    idx = -1
    for idx, layout_shape in enumerate(layout_shapes):
        assert isinstance(layout_shape, BaseShape)
    assert idx == 2


@then('iterating master_placeholders produces 2 MasterPlaceholder objects')
def then_iterating_master_placeholders_produces_2_objects(context):
    master_placeholders = context.master_placeholders
    idx = -1
    for idx, master_placeholder in enumerate(master_placeholders):
        assert type(master_placeholder).__name__ == 'MasterPlaceholder'
    assert idx == 1


@then('iterating master_shapes produces 2 BaseShape objects')
def then_iterating_master_shapes_produces_2_BaseShape_objects(context):
    master_shapes = context.master_shapes
    idx = -1
    for idx, master_shape in enumerate(master_shapes):
        assert isinstance(master_shape, BaseShape)
    assert idx == 1


@then('iterating placeholders produces 2 SlidePlaceholder objects')
def then_iterating_placeholders_produces_2_SlidePlaceholder_objects(context):
    placeholders = context.placeholders
    idx = -1
    for idx, placeholder in enumerate(placeholders):
        assert type(placeholder).__name__ == 'SlidePlaceholder'
    assert idx == 1


@then('iterating produces 3 NotesSlidePlaceholder objects')
def then_iterating_produces_3_NotesSlidePlaceholder_objects(context):
    idx = -1
    for idx, placeholder in enumerate(context.notes_slide.placeholders):
        typename = type(placeholder).__name__
        assert typename == 'NotesSlidePlaceholder', 'got %s' % typename
    assert idx == 2


@then('iterating shapes produces 6 BaseShape objects')
def then_iterating_shapes_produces_6_BaseShape_objects(context):
    shapes = context.shapes
    idx = -1
    for idx, shape in enumerate(shapes):
        assert isinstance(shape, BaseShape)
    assert idx == 5


@then('iterating slide_layouts produces 2 SlideLayout objects')
def then_iterating_slide_layouts_produces_2_SlideLayout_objects(context):
    slide_layouts = context.slide_layouts
    idx = -1
    for idx, slide_layout in enumerate(slide_layouts):
        assert type(slide_layout).__name__ == 'SlideLayout'
    assert idx == 1


@then('iterating slide_masters produces 2 SlideMaster objects')
def then_iterating_slide_masters_produces_2_SlideMaster_objects(context):
    slide_masters = context.slide_masters
    idx = -1
    for idx, slide_master in enumerate(slide_masters):
        assert type(slide_master).__name__ == 'SlideMaster'
    assert idx == 1


@then('iterating slides produces 3 Slide objects')
def then_iterating_slides_produces_3_Slide_objects(context):
    slides = context.slides
    idx = -1
    for idx, slide in enumerate(slides):
        assert type(slide).__name__ == 'Slide'
    assert idx == 2


@then('layout_placeholders[1] is a LayoutPlaceholder object')
def then_layout_placeholders_1_is_a_LayoutPlaceholder_object(context):
    layout_placeholders = context.layout_placeholders
    assert type(layout_placeholders[1]).__name__ == 'LayoutPlaceholder'


@then('layout_placeholders.get(idx=10) is the body placeholder')
def then_layout_placeholders_get_10_is_the_body_placeholder(context):
    layout_placeholders = context.layout_placeholders
    title_placeholder = layout_placeholders.get(idx=0)
    body_placeholder = layout_placeholders.get(idx=10)
    assert title_placeholder._element is layout_placeholders[0]._element
    assert body_placeholder._element is layout_placeholders[1]._element


@then('layout_shapes[1] is a LayoutPlaceholder object')
def then_layout_shapes_1_is_a_LayoutPlaceholder_object(context):
    layout_shapes = context.layout_shapes
    assert type(layout_shapes[1]).__name__ == 'LayoutPlaceholder'


@then('len(layout_placeholders) is 2')
def then_len_layout_placeholders_is_2(context):
    layout_placeholders = context.layout_placeholders
    assert len(layout_placeholders) == 2


@then('len(layout_shapes) is 3')
def then_len_layout_shapes_is_3(context):
    layout_shapes = context.layout_shapes
    assert len(layout_shapes) == 3


@then('len(master_placeholders) is 2')
def then_len_master_placeholders_is_2(context):
    master_placeholders = context.master_placeholders
    assert len(master_placeholders) == 2


@then('len(master_shapes) is 2')
def then_len_master_shapes_is_2(context):
    master_shapes = context.master_shapes
    assert len(master_shapes) == 2


@then('len(notes_slide.shapes) is {count}')
def then_len_notes_slide_shapes_is_count(context, count):
    shapes = context.notes_slide.shapes
    assert len(shapes) == int(count)


@then('len(placeholders) is 2')
def then_len_placeholders_is_2(context):
    placeholders = context.placeholders
    assert len(placeholders) == 2


@then('len(shapes) is 6')
def then_len_shapes_is_6(context):
    shapes = context.shapes
    assert len(shapes) == 6


@then('len(slides) is {count}')
def then_len_slides_is_count(context, count):
    slides = context.slides
    assert len(slides) == int(count)


@then('len(slide_layouts) is 2')
def then_len_slide_layouts_is_2(context):
    slide_layouts = context.slide_layouts
    assert len(slide_layouts) == 2


@then('len(slide_masters) is 2')
def then_len_slide_masters_is_2(context):
    slide_masters = context.slide_masters
    assert len(slide_masters) == 2


@then('master_placeholders[1] is a MasterPlaceholder object')
def then_master_placeholders_1_is_a_MasterPlaceholder_object(context):
    master_placeholders = context.master_placeholders
    assert type(master_placeholders[1]).__name__ == 'MasterPlaceholder'


@then('master_placeholders.get(PP_PLACEHOLDER.BODY) is the body placeholder')
def then_master_placeholders_get_PP_PLACEHOLDER_BODY_is_the_body_ph(context):
    master_placeholders = context.master_placeholders
    title_placeholder = master_placeholders.get(PP_PLACEHOLDER.TITLE)
    body_placeholder = master_placeholders.get(PP_PLACEHOLDER.BODY)
    assert title_placeholder._element is master_placeholders[0]._element
    assert body_placeholder._element is master_placeholders[1]._element


@then('master_shapes[1] is a Picture object')
def then_master_shapes_1_is_a_Picture_object(context):
    master_shapes = context.master_shapes
    assert type(master_shapes[1]).__name__ == 'Picture'


@then('notes_slide.notes_placeholder is a NotesSlidePlaceholder object')
def then_notes_slide_notes_placeholder_is_a_NotesSlidePlacehldr_obj(context):
    notes_slide = context.notes_slide
    cls_name = type(notes_slide.notes_placeholder).__name__
    assert cls_name == 'NotesSlidePlaceholder', 'got %s' % cls_name


@then('notes_slide.notes_text_frame is a TextFrame object')
def then_notes_slide_notes_text_frame_is_a_TextFrame_object(context):
    notes_slide = context.notes_slide
    cls_name = type(notes_slide.notes_text_frame).__name__
    assert cls_name == 'TextFrame', 'got %s' % cls_name


@then('notes_slide.placeholders is a NotesSlidePlaceholders object')
def then_notes_slide_placeholders_is_a_NotesSlidePlaceholders_object(context):
    notes_slide = context.notes_slide
    assert type(notes_slide.placeholders).__name__ == 'NotesSlidePlaceholders'


@then('placeholders[10] is a SlidePlaceholder object')
def then_placeholders_10_is_a_SlidePlaceholder_object(context):
    placeholders = context.placeholders
    assert type(placeholders[10]).__name__ == 'SlidePlaceholder'


@then('shapes[{idx}] is a {type_} object')
def then_shapes_idx_is_a_type_object(context, idx, type_):
    shapes = context.shapes
    type_name = type(shapes[int(idx)]).__name__
    assert type_name == type_, 'got %s' % type_name


@then('shapes.title is the title placeholder')
def then_shapes_title_is_the_title_placeholder(context):
    shapes = context.shapes
    title_placeholder = shapes.title
    assert title_placeholder.element is shapes[0].element
    assert title_placeholder.shape_id == 4


@then('slide.has_notes_slide is {value}')
def then_slide_has_notes_slide_is_value(context, value):
    expected_value = {'True': True, 'False': False}[value]
    slide = context.slide
    assert slide.has_notes_slide is expected_value


@then('slide.name is {value}')
def then_slide_name_is_value(context, value):
    expected_value = '' if value == 'the empty string' else value
    slide = context.slide
    assert slide.name == expected_value


@then('slide.notes_slide is a NotesSlide object')
def then_slide_notes_slide_is_a_NotesSlide_object(context):
    notes_slide = context.notes_slide = context.slide.notes_slide
    assert type(notes_slide).__name__ == 'NotesSlide'


@then('slide.placeholders is a SlidePlaceholders object')
def then_slide_placeholders_is_a_SlidePlaceholders_object(context):
    slide = context.slide
    assert type(slide.placeholders).__name__ == 'SlidePlaceholders'


@then('slide.shapes is a SlideShapes object')
def then_slide_shapes_is_a_SlideShapes_object(context):
    slide = context.slide
    assert type(slide.shapes).__name__ == 'SlideShapes'


@then('slide.slide_id is 256')
def then_slide_slide_id_is_256(context):
    slide = context.slide
    assert slide.slide_id == 256


@then('slide.slide_layout is the one passed in the call')
def then_slide_slide_layout_is_the_one_passed_in_the_call(context):
    slide = context.prs.slides[3]
    assert slide.slide_layout == context.slide_layout


@then('slide_layout.name is {value}')
def then_slide_layout_name_is_value(context, value):
    expected_value = '' if value == 'the empty string' else value
    slide_layout = context.slide_layout
    assert slide_layout.name == expected_value, 'got %s' % slide_layout.name


@then('slide_layout.placeholders is a LayoutPlaceholders object')
def then_slide_layout_placeholders_is_a_LayoutPlaceholders_object(context):
    slide_layout = context.slide_layout
    assert type(slide_layout.placeholders).__name__ == 'LayoutPlaceholders'


@then('slide_layout.shapes is a LayoutShapes object')
def then_slide_layout_shapes_is_a_LayoutShapes_object(context):
    slide_layout = context.slide_layout
    assert type(slide_layout.shapes).__name__ == 'LayoutShapes'


@then('slide_layout.slide_master is a SlideMaster object')
def then_slide_layout_slide_master_is_a_SlideMaster_object(context):
    slide_layout = context.slide_layout
    assert type(slide_layout.slide_master).__name__ == 'SlideMaster'


@then('slide_layouts[1] is a SlideLayout object')
def then_slide_layouts_1_is_a_SlideLayout_object(context):
    slide_layouts = context.slide_layouts
    assert type(slide_layouts[1]).__name__ == 'SlideLayout'


@then('slide_master.placeholders is a MasterPlaceholders object')
def then_slide_master_placeholders_is_a_MasterPlaceholders_object(context):
    slide_master = context.slide_master
    assert type(slide_master.placeholders).__name__ == 'MasterPlaceholders'


@then('slide_master.shapes is a MasterShapes object')
def then_slide_master_shapes_is_a_MasterShapes_object(context):
    slide_master = context.slide_master
    assert type(slide_master.shapes).__name__ == 'MasterShapes'


@then('slide_master.slide_layouts is a SlideLayouts object')
def then_slide_master_slide_layouts_is_a_SlideLayouts_object(context):
    slide_master = context.slide_master
    assert type(slide_master.slide_layouts).__name__ == 'SlideLayouts'


@then('slide_masters[1] is a SlideMaster object')
def then_slide_masters_1_is_a_SlideMaster_object(context):
    slide_masters = context.slide_masters
    assert type(slide_masters[1]).__name__ == 'SlideMaster'


@then('slides.get(256) is slides[0]')
def then_slides_get_256_is_slides_0(context):
    slides = context.slides
    assert slides.get(256) is slides[0]


@then('slides.get(666, default=slides[2]) is slides[2]')
def then_slides_get_666_default_slides_2_is_slides_2(context):
    slides = context.slides
    assert slides.get(666, default=slides[2]) is slides[2]


@then('slides[2] is a Slide object')
def then_slides_2_is_a_Slide_object(context):
    slides = context.slides
    assert type(slides[2]).__name__ == 'Slide'


@then('the index of each shape matches its position in the sequence')
def then_index_of_each_shape_matches_its_position_in_the_sequence(context):
    shapes = context.shapes
    for idx, shape in enumerate(shapes):
        assert idx == shapes.index(shape), (
            "index doesn't match for idx == %s" % idx
        )
