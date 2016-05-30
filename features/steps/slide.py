# encoding: utf-8

"""
Gherkin step implementations for slide-related features.
"""

from __future__ import absolute_import

from behave import given, when, then

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.shapes.base import BaseShape
from pptx.shapes.factory import SlidePlaceholders
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.picture import Picture
from pptx.shapes.placeholder import (
    LayoutPlaceholder, MasterPlaceholder, SlidePlaceholder
)
from pptx.shapes.shapetree import (
    LayoutPlaceholders, LayoutShapes, MasterPlaceholders, MasterShapes,
    SlideShapeTree
)
from pptx.slide import SlideLayouts

from helpers import test_pptx

SHAPE_COUNT = 3


# given ===================================================

@given('a blank slide')
def given_a_blank_slide(context):
    context.prs = Presentation(test_pptx('sld-blank'))
    context.slide = context.prs.slides[0]


@given('a layout placeholder collection')
def given_layout_placeholder_collection(context):
    prs = Presentation(test_pptx('lyt-shapes'))
    context.layout_placeholders = prs.slide_layouts[0].placeholders


@given('a layout shape collection')
def given_layout_shape_collection(context):
    prs = Presentation(test_pptx('lyt-shapes'))
    context.layout_shapes = prs.slide_layouts[0].shapes


@given('a master placeholder collection')
def given_master_placeholder_collection(context):
    prs = Presentation(test_pptx('mst-placeholders'))
    context.master_placeholders = prs.slide_master.placeholders


@given('a master shape collection containing two shapes')
def given_master_shape_collection_containing_two_shapes(context):
    prs = Presentation(test_pptx('mst-shapes'))
    context.master_shapes = prs.slide_master.shapes


@given('a slide having a title')
def given_a_slide_having_a_title(context):
    prs = Presentation(test_pptx('sld-access-shapes'))
    context.prs, context.slide = prs, prs.slides[0]


@given('a slide having six shapes')
def given_a_slide_having_six_shapes(context):
    presentation = Presentation(test_pptx('sld-access-shapes'))
    context.slide = presentation.slides[0]


@given('a slide having two placeholders')
def given_a_slide_having_two_placeholders(context):
    prs = Presentation(test_pptx('sld-access-shapes'))
    context.slide = prs.slides[0]


@given('a slide layout')
def given_a_slide_layout(context):
    prs = Presentation(test_pptx('sld-slide-access'))
    context.slide_layout = prs.slide_layouts[0]


@given('a slide layout collection containing two layouts')
def given_slide_layout_collection_containing_two_layouts(context):
    prs = Presentation(test_pptx('mst-slide-layouts'))
    context.slide_layouts = prs.slide_master.slide_layouts


@given('a slide layout having three shapes')
def given_slide_layout_having_three_shapes(context):
    prs = Presentation(test_pptx('lyt-shapes'))
    context.slide_layout = prs.slide_layouts[0]


@given('a slide layout having two placeholders')
def given_layout_having_two_placeholders(context):
    prs = Presentation(test_pptx('lyt-shapes'))
    context.slide_layout = prs.slide_layouts[0]


@given('a slide master')
def given_a_slide_master(context):
    prs = Presentation(test_pptx('sld-slide-access'))
    context.slide_master = prs.slide_masters[0]


@given('a slide master having two placeholders')
def given_master_having_two_placeholders(context):
    prs = Presentation(test_pptx('mst-placeholders'))
    context.slide_master = prs.slide_master


@given('a slide master having two shapes')
def given_slide_master_having_two_shapes(context):
    prs = Presentation(test_pptx('mst-shapes'))
    context.slide_master = prs.slide_master


@given('a slide master having two slide layouts')
def given_slide_master_having_two_layouts(context):
    prs = Presentation(test_pptx('mst-slide-layouts'))
    context.slide_master = prs.slide_master


@given('a slide placeholder collection')
def given_a_slide_placeholder_collection(context):
    prs = Presentation(test_pptx('sld-access-shapes'))
    context.slide_placeholders = prs.slides[0].placeholders


@given('a slide shape collection')
def given_a_slide_shape_collection(context):
    presentation = Presentation(test_pptx('sld-access-shapes'))
    context.shapes = presentation.slides[0].shapes


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

@then('each shape is of the appropriate type')
def then_each_shape_is_of_appropriate_type(context):
    layout_shapes = context.layout_shapes
    expected_types = [LayoutPlaceholder, LayoutPlaceholder, Picture]
    for idx, layout_shape in enumerate(layout_shapes):
        assert type(layout_shape) == expected_types[idx], (
            "got \'%s\'" % type(layout_shape).__name__
        )


@then('each slide shape is of the appropriate type')
def then_each_slide_shape_is_of_the_appropriate_type(context):

    def assertShapeType(shape, expected_type):
        assert type(shape) is expected_type, (
            "got \'%s\'" % type(shape).__name__
        )

    shapes = context.shapes
    assertShapeType(shapes[0], SlidePlaceholder)
    assertShapeType(shapes[1], SlidePlaceholder)
    assertShapeType(shapes[2], Picture)
    assertShapeType(shapes[3], GraphicFrame)
    assertShapeType(shapes[4], GraphicFrame)
    assertShapeType(shapes[5], GraphicFrame)


@then('I can access a layout placeholder by idx value')
def then_can_access_layout_placeholder_by_idx_value(context):
    layout_placeholders = context.layout_placeholders
    title_placeholder = layout_placeholders.get(idx=0)
    body_placeholder = layout_placeholders.get(idx=10)
    assert title_placeholder._element is layout_placeholders[0]._element
    assert body_placeholder._element is layout_placeholders[1]._element


@then('I can access a layout placeholder by index')
def then_can_access_layout_placeholder_by_index(context):
    layout_placeholders = context.layout_placeholders
    for idx in range(2):
        layout_placeholder = layout_placeholders[idx]
        assert isinstance(layout_placeholder, LayoutPlaceholder)


@then('I can access a layout shape by index')
def then_can_access_layout_shape_by_index(context):
    layout_shapes = context.layout_shapes
    for idx in range(SHAPE_COUNT):
        layout_shape = layout_shapes[idx]
        assert isinstance(layout_shape, BaseShape)


@then('I can access a master placeholder by index')
def then_can_access_master_placeholder_by_index(context):
    master_placeholders = context.master_placeholders
    for idx in range(2):
        master_placeholder = master_placeholders[idx]
        assert isinstance(master_placeholder, MasterPlaceholder)


@then('I can access a master placeholder by type')
def then_can_access_master_placeholder_by_type(context):
    master_placeholders = context.master_placeholders
    title_placeholder = master_placeholders.get(PP_PLACEHOLDER.TITLE)
    body_placeholder = master_placeholders.get(PP_PLACEHOLDER.BODY)
    assert title_placeholder._element is master_placeholders[0]._element
    assert body_placeholder._element is master_placeholders[1]._element


@then('I can access a master shape by index')
def then_can_access_master_shape_by_index(context):
    master_shapes = context.master_shapes
    for idx in range(2):
        master_shape = master_shapes[idx]
        assert isinstance(master_shape, BaseShape)


@then('I can access a shape by index')
def then_can_access_shape_by_index(context):
    shapes = context.shapes
    for idx in range(2):
        shape = shapes[idx]
        assert isinstance(shape, BaseShape)


@then('I can access a slide layout by index')
def then_can_access_slide_layout_by_index(context):
    slide_layouts = context.slide_layouts
    for idx in range(2):
        assert type(slide_layouts[idx]).__name__ == 'SlideLayout'


@then('I can access a slide placeholder by index')
def then_can_access_slide_placeholder_by_index(context):
    slide_placeholders = context.slide_placeholders
    for idx in (0, 10):
        slide_placeholder = slide_placeholders[idx]
        assert isinstance(slide_placeholder, SlidePlaceholder)


@then('I can access the placeholder collection of the slide')
def then_can_access_placeholder_collection_of_slide(context):
    slide = context.slide
    slide_placeholders = slide.placeholders
    msg = 'Slide.placeholders not instance of SlidePlaceholders'
    assert isinstance(slide_placeholders, SlidePlaceholders), msg


@then('I can access the placeholder collection of the slide layout')
def then_can_access_placeholder_collection_of_slide_layout(context):
    slide_layout = context.slide_layout
    layout_placeholders = slide_layout.placeholders
    msg = 'SlideLayout.placeholders not instance of LayoutPlaceholders'
    assert isinstance(layout_placeholders, LayoutPlaceholders), msg


@then('I can access the placeholder collection of the slide master')
def then_can_access_placeholder_collection_of_slide_master(context):
    slide_master = context.slide_master
    master_placeholders = slide_master.placeholders
    msg = 'SlideMaster.placeholders not instance of MasterPlaceholders'
    assert isinstance(master_placeholders, MasterPlaceholders), msg


@then('I can access the shape collection of the slide')
def then_can_access_shapes_of_slide(context):
    slide = context.slide
    shapes = slide.shapes
    msg = 'Slide.shapes not instance of SlideShapeTree'
    assert isinstance(shapes, SlideShapeTree), msg


@then('I can access the shape collection of the slide layout')
def then_can_access_shape_collection_of_slide_layout(context):
    slide_layout = context.slide_layout
    layout_shapes = slide_layout.shapes
    msg = 'SlideLayout.shapes not instance of LayoutShapes'
    assert isinstance(layout_shapes, LayoutShapes), msg


@then('I can access the shape collection of the slide master')
def then_can_access_shape_collection_of_slide_master(context):
    slide_master = context.slide_master
    master_shapes = slide_master.shapes
    msg = 'SlideMaster.shapes not instance of MasterShapes'
    assert isinstance(master_shapes, MasterShapes), msg


@then('I can access the slide layouts of the slide master')
def then_can_access_slide_layouts_of_slide_master(context):
    slide_master = context.slide_master
    slide_layouts = slide_master.slide_layouts
    msg = 'SlideMaster.slide_layouts not instance of SlideLayouts'
    assert isinstance(slide_layouts, SlideLayouts), msg


@then('I can access the title placeholder')
def then_I_can_access_the_title_placeholder(context):
    shapes = context.shapes
    title_placeholder = shapes.title
    assert title_placeholder.element is shapes[0].element
    assert title_placeholder.id == 4


@then('I can iterate over the layout placeholders')
def then_can_iterate_over_the_layout_placeholders(context):
    layout_placeholders = context.layout_placeholders
    actual_count = 0
    for layout_placeholder in layout_placeholders:
        actual_count += 1
        assert isinstance(layout_placeholder, LayoutPlaceholder)
    assert actual_count == 2


@then('I can iterate over the layout shapes')
def then_can_iterate_over_the_layout_shapes(context):
    layout_shapes = context.layout_shapes
    actual_count = 0
    for layout_shape in layout_shapes:
        actual_count += 1
        assert isinstance(layout_shape, BaseShape)
    assert actual_count == SHAPE_COUNT


@then('I can iterate over the master placeholders')
def then_can_iterate_over_the_master_placeholders(context):
    master_placeholders = context.master_placeholders
    actual_count = 0
    for master_placeholder in master_placeholders:
        actual_count += 1
        assert isinstance(master_placeholder, MasterPlaceholder)
    assert actual_count == 2


@then('I can iterate over the master shapes')
def then_can_iterate_over_the_master_shapes(context):
    master_shapes = context.master_shapes
    actual_count = 0
    for master_shape in master_shapes:
        actual_count += 1
        assert isinstance(master_shape, BaseShape)
    assert actual_count == 2


@then('I can iterate over the shapes')
def then_can_iterate_over_the_shapes(context):
    shapes = context.shapes
    actual_count = 0
    for shape in shapes:
        actual_count += 1
        assert isinstance(shape, BaseShape)
    assert actual_count == 6


@then('I can iterate over the slide placeholders')
def then_can_iterate_over_the_slide_placeholders(context):
    slide_placeholders = context.slide_placeholders
    actual_count = 0
    for slide_placeholder in slide_placeholders:
        actual_count += 1
        assert isinstance(slide_placeholder, SlidePlaceholder)
    assert actual_count == 2


@then('I can iterate slide_layouts')
def then_I_can_iterate_slide_layouts(context):
    slide_layouts = context.slide_layouts
    idx = 0
    for idx, slide_layout in enumerate(slide_layouts):
        assert type(slide_layout).__name__ == 'SlideLayout'
    assert idx == 1


@then('iterating slides produces 3 Slide objects')
def then_iterating_slides_produces_3_slide_objects(context):
    slides = context.slides
    idx = -1
    for idx, slide in enumerate(slides):
        assert type(slide).__name__ == 'Slide'
    assert idx == 2


@then('len(slides) is {count}')
def then_len_slides_is_count(context, count):
    slides = context.slides
    assert len(slides) == int(count)


@then('len(slide_layouts) is 2')
def then_len_slide_layouts_is_2(context):
    slide_master = context.slide_master
    slide_layouts = slide_master.slide_layouts
    assert len(slide_layouts) == 2, (
        'expected len(slide_layouts) of 2, got %s' % len(slide_layouts)
    )


@then('slide.slide_layout is the one passed in the call')
def then_slide_slide_layout_is_the_one_passed_in_the_call(context):
    slide = context.prs.slides[3]
    assert slide.slide_layout == context.slide_layout


@then('slides[2] is a Slide object')
def then_slides_2_is_a_Slide_object(context):
    slides = context.slides
    assert type(slides[2]).__name__ == 'Slide'


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


@then('the index of each shape matches its position in the sequence')
def then_index_of_each_shape_matches_its_position_in_the_sequence(context):
    shapes = context.shapes
    for idx, shape in enumerate(shapes):
        assert idx == shapes.index(shape), (
            "index doesn't match for idx == %s" % idx
        )


@then('the length of the layout placeholder collection is 2')
def then_len_of_placeholder_collection_is_2(context):
    slide_layout = context.slide_layout
    layout_placeholders = slide_layout.placeholders
    assert len(layout_placeholders) == 2, (
        'expected len(layout_placeholders) of 2, got %s' %
        len(layout_placeholders)
    )


@then('the length of the layout shape collection counts all its shapes')
def then_len_of_layout_shape_collection_counts_all_its_shapes(context):
    slide_layout = context.slide_layout
    layout_shapes = slide_layout.shapes
    assert len(layout_shapes) == SHAPE_COUNT, (
        'expected len(layout_shapes) of %d, got %s' %
        (SHAPE_COUNT, len(layout_shapes))
    )


@then('the length of the master placeholder collection is 2')
def then_len_of_master_placeholder_collection_is_2(context):
    slide_master = context.slide_master
    master_placeholders = slide_master.placeholders
    assert len(master_placeholders) == 2, (
        'expected len(master_placeholders) of 2, got %s' %
        len(master_placeholders)
    )


@then('the length of the master shape collection is 2')
def then_len_of_master_shape_collection_is_2(context):
    slide_master = context.slide_master
    master_shapes = slide_master.shapes
    assert len(master_shapes) == 2, (
        'expected len(master_shapes) of 2, got %s' % len(master_shapes)
    )


@then('the length of the shape collection is 6')
def then_len_of_shape_collection_is_6(context):
    slide = context.slide
    shapes = slide.shapes
    assert len(shapes) == 6, (
        'expected len(shapes) of 6, got %s' % len(shapes)
    )


@then('the length of the slide placeholder collection is 2')
def then_len_of_slide_placeholder_collection_is_2(context):
    slide = context.slide
    slide_placeholders = slide.placeholders
    assert len(slide_placeholders) == 2, (
        'expected len(slide_placeholders) of 2, got %s' %
        len(slide_placeholders)
    )
