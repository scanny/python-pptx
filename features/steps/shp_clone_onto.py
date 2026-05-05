"""Gherkin step implementations for `BaseShape.clone_onto`."""

from __future__ import annotations

from behave import given, then, when
from helpers import test_image

from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR_TYPE, MSO_SHAPE
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.group import GroupShape
from pptx.shapes.picture import Picture
from pptx.util import Inches

# given ===================================================


@given("a slide with an auto-shape")
def given_a_slide_with_an_autoshape(context):
    prs = Presentation()
    context.prs = prs
    context.src_slide = prs.slides.add_slide(prs.slide_layouts[5])
    rect = context.src_slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
    )
    rect.text_frame.text = "Greetings"
    context.src_shape = rect


@given("a slide with a picture")
def given_a_slide_with_a_picture(context):
    prs = Presentation()
    context.prs = prs
    context.src_slide = prs.slides.add_slide(prs.slide_layouts[5])
    context.src_shape = context.src_slide.shapes.add_picture(
        test_image("python-icon.jpeg"), Inches(1), Inches(1)
    )


@given("a slide with a connector")
def given_a_slide_with_a_connector(context):
    prs = Presentation()
    context.prs = prs
    context.src_slide = prs.slides.add_slide(prs.slide_layouts[5])
    context.src_shape = context.src_slide.shapes.add_connector(
        MSO_CONNECTOR_TYPE.STRAIGHT, Inches(1), Inches(1), Inches(3), Inches(1)
    )


@given("a slide with a table")
def given_a_slide_with_a_table(context):
    prs = Presentation()
    context.prs = prs
    context.src_slide = prs.slides.add_slide(prs.slide_layouts[5])
    context.src_shape = context.src_slide.shapes.add_table(
        2, 2, Inches(1), Inches(1), Inches(3), Inches(1)
    )


@given("a slide with a group shape")
def given_a_slide_with_a_group(context):
    prs = Presentation()
    context.prs = prs
    context.src_slide = prs.slides.add_slide(prs.slide_layouts[5])
    context.src_shape = context.src_slide.shapes.add_group_shape()


@given("a second slide in the same presentation")
def given_a_second_slide_same_prs(context):
    context.tgt_slide = context.prs.slides.add_slide(context.prs.slide_layouts[5])


@given("a source presentation with a picture on slide 1")
def given_source_prs_with_pic(context):
    prs = Presentation()
    context.src_prs = prs
    context.src_slide = prs.slides.add_slide(prs.slide_layouts[5])
    context.src_shape = context.src_slide.shapes.add_picture(
        test_image("python-icon.jpeg"), Inches(1), Inches(1)
    )


@given("a target presentation with one empty slide")
def given_target_prs(context):
    prs = Presentation()
    context.tgt_prs = prs
    context.tgt_slide = prs.slides.add_slide(prs.slide_layouts[5])


# when ====================================================


@when("I clone the auto-shape onto the second slide")
def when_clone_autoshape(context):
    context.clone = context.src_shape.clone_onto(context.tgt_slide.shapes)


@when("I clone the picture onto the second slide")
def when_clone_picture(context):
    context.clone = context.src_shape.clone_onto(context.tgt_slide.shapes)


@when("I clone the picture onto the target presentation's slide")
def when_clone_picture_cross_prs(context):
    context.clone = context.src_shape.clone_onto(context.tgt_slide.shapes)


@when("I clone the connector onto the second slide")
def when_clone_connector(context):
    context.clone = context.src_shape.clone_onto(context.tgt_slide.shapes)


@when("I clone the table onto the second slide")
def when_clone_table(context):
    context.clone = context.src_shape.clone_onto(context.tgt_slide.shapes)


@when("I clone the group onto the second slide")
def when_clone_group(context):
    context.clone = context.src_shape.clone_onto(context.tgt_slide.shapes)


@when("I clone the auto-shape onto the second slide at (3in, 4in)")
def when_clone_with_pos(context):
    context.clone = context.src_shape.clone_onto(
        context.tgt_slide.shapes, left=Inches(3), top=Inches(4)
    )


# then ====================================================


@then("the clone is a Shape with a fresh shape id")
def then_clone_is_shape_fresh_id(context):
    from pptx.shapes.autoshape import Shape

    assert isinstance(context.clone, Shape), "clone is %r" % type(context.clone)
    assert context.clone.shape_id != context.src_shape.shape_id


@then("the clone has the same text as the source")
def then_clone_same_text(context):
    assert context.clone.text_frame.text == context.src_shape.text_frame.text


@then("the clone is topmost in the second slide's z-order")
def then_clone_topmost(context):
    assert context.tgt_slide.shapes[-1]._element is context.clone._element


@then("the clone is a Picture")
def then_clone_is_picture(context):
    assert isinstance(context.clone, Picture), "clone is %r" % type(context.clone)


@then("the clone's image bytes match the source")
def then_clone_image_bytes(context):
    assert context.clone.image is not None
    assert context.clone.image.blob == context.src_shape.image.blob


@then("the clone's image part lives in the target presentation's package")
def then_clone_image_in_tgt_pkg(context):
    assert context.clone.part.package is context.tgt_prs.part.package


@then("the clone is a Connector with the same begin and end points")
def then_clone_connector(context):
    from pptx.shapes.connector import Connector

    assert isinstance(context.clone, Connector)
    assert context.clone.begin_x == context.src_shape.begin_x
    assert context.clone.end_x == context.src_shape.end_x
    assert context.clone.begin_y == context.src_shape.begin_y
    assert context.clone.end_y == context.src_shape.end_y


@then("the clone is a GraphicFrame that has a table")
def then_clone_gf_table(context):
    assert isinstance(context.clone, GraphicFrame)
    assert context.clone.has_table


@then("the clone is a GroupShape with a fresh shape id")
def then_clone_group_fresh_id(context):
    assert isinstance(context.clone, GroupShape)
    assert context.clone.shape_id != context.src_shape.shape_id


@then("the clone's position is (3in, 4in)")
def then_clone_position(context):
    assert context.clone.left == Inches(3)
    assert context.clone.top == Inches(4)


# ---------- CLO-11: group-clone regression scenarios ---------------------


@given("a slide with a group of three autoshapes")
def given_a_slide_with_a_group_of_three(context):
    prs = Presentation()
    context.prs = prs
    context.src_slide = prs.slides.add_slide(prs.slide_layouts[5])
    s1 = context.src_slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(1), Inches(1)
    )
    s2 = context.src_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(2), Inches(2), Inches(1), Inches(1)
    )
    s3 = context.src_slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3), Inches(3), Inches(1), Inches(1)
    )
    context.src_shape = context.src_slide.shapes.add_group_shape([s1, s2, s3])
    context.src_child_positions = [
        (c.left, c.top, c.width, c.height) for c in context.src_shape.shapes
    ]


@given("a slide with a nested group shape")
def given_a_slide_with_a_nested_group(context):
    prs = Presentation()
    context.prs = prs
    context.src_slide = prs.slides.add_slide(prs.slide_layouts[5])
    s1 = context.src_slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(1), Inches(1)
    )
    s2 = context.src_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(2), Inches(2), Inches(1), Inches(1)
    )
    inner = context.src_slide.shapes.add_group_shape([s1, s2])
    s3 = context.src_slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4), Inches(4), Inches(1), Inches(1)
    )
    context.src_shape = context.src_slide.shapes.add_group_shape([inner, s3])


@given("a slide with a group containing a picture")
def given_a_slide_with_group_containing_picture(context):
    prs = Presentation()
    context.prs = prs
    context.src_slide = prs.slides.add_slide(prs.slide_layouts[5])
    pic = context.src_slide.shapes.add_picture(
        test_image("python-icon.jpeg"), Inches(1), Inches(1)
    )
    shp = context.src_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(3), Inches(1), Inches(1), Inches(1)
    )
    context.src_picture = pic
    context.src_shape = context.src_slide.shapes.add_group_shape([pic, shp])


@given("a slide with a group containing a chart")
def given_a_slide_with_group_containing_chart(context):
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    prs = Presentation()
    context.prs = prs
    context.src_slide = prs.slides.add_slide(prs.slide_layouts[5])
    cd = CategoryChartData()
    cd.categories = ["A", "B", "C"]
    cd.add_series("Series 1", (1, 2, 3))
    chart_gf = context.src_slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(4),
        Inches(3),
        cd,
    )
    shp = context.src_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(5), Inches(1), Inches(1), Inches(1)
    )
    context.src_shape = context.src_slide.shapes.add_group_shape([chart_gf, shp])


@given("a slide with a group of two autoshapes")
def given_a_slide_with_group_of_two(context):
    prs = Presentation()
    context.prs = prs
    context.src_slide = prs.slides.add_slide(prs.slide_layouts[5])
    s1 = context.src_slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(1), Inches(1)
    )
    s2 = context.src_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(2), Inches(2), Inches(1), Inches(1)
    )
    context.src_shape = context.src_slide.shapes.add_group_shape([s1, s2])


@when("I clone the nested group onto the second slide")
def when_clone_nested_group(context):
    context.clone = context.src_shape.clone_onto(context.tgt_slide.shapes)


@when("I clone the group onto the second slide at (5in, 5in)")
def when_clone_group_with_pos(context):
    context.clone = context.src_shape.clone_onto(
        context.tgt_slide.shapes, left=Inches(5), top=Inches(5)
    )


@then("the clone is a GroupShape containing three autoshapes")
def then_clone_is_group_of_three(context):
    assert isinstance(context.clone, GroupShape)
    assert len(context.clone.shapes) == 3


@then("the clone's children have the same relative positions")
def then_clone_children_same_positions(context):
    clone_positions = [(c.left, c.top, c.width, c.height) for c in context.clone.shapes]
    assert clone_positions == context.src_child_positions


@then("every descendant shape survives on the clone")
def then_every_descendant_survives(context):
    # -- outer(1) + inner(1) + three autoshapes (3) = 5 cNvPrs in total --
    assert len(context.clone._element.xpath(".//p:cNvPr")) == 5
    # -- inner nested group still has its two children --
    inner_clones = [s for s in context.clone.shapes if isinstance(s, GroupShape)]
    assert len(inner_clones) == 1
    assert len(inner_clones[0].shapes) == 2


@then("every descendant id is unique and fresh")
def then_every_descendant_id_fresh(context):
    clone_ids = [
        int(cnv.get("id")) for cnv in context.clone._element.xpath(".//p:cNvPr")
    ]
    src_ids = [
        int(cnv.get("id")) for cnv in context.src_shape._element.xpath(".//p:cNvPr")
    ]
    assert not (set(clone_ids) & set(src_ids))
    assert len(set(clone_ids)) == len(clone_ids)


@then("the cloned group contains a picture with matching image bytes")
def then_cloned_group_picture(context):
    pic_clones = [s for s in context.clone.shapes if isinstance(s, Picture)]
    assert len(pic_clones) == 1
    assert pic_clones[0].image is not None
    assert pic_clones[0].image.blob == context.src_picture.image.blob


@then("the cloned group contains a chart resolvable on the target part")
def then_cloned_group_chart(context):
    chart_in_clone = None
    for s in context.clone.shapes:
        if isinstance(s, GraphicFrame) and s.has_chart:
            chart_in_clone = s
            break
    assert chart_in_clone is not None
    rIds = chart_in_clone._element.xpath(".//@r:id")
    assert rIds
    tgt_rels = context.tgt_slide.part.rels
    for rId in rIds:
        assert rId in tgt_rels


@then("the cloned group is at (5in, 5in)")
def then_cloned_group_at_5_5(context):
    assert context.clone.left == Inches(5)
    assert context.clone.top == Inches(5)


@then("the cloned group's child coordinate system is unchanged")
def then_child_coord_system_unchanged(context):
    src_xfrm = context.src_shape._element.xfrm
    clone_xfrm = context.clone._element.xfrm
    assert src_xfrm is not None
    assert clone_xfrm is not None
    assert clone_xfrm.xpath("./a:chOff/@x") == src_xfrm.xpath("./a:chOff/@x")
    assert clone_xfrm.xpath("./a:chOff/@y") == src_xfrm.xpath("./a:chOff/@y")
    assert clone_xfrm.xpath("./a:chExt/@cx") == src_xfrm.xpath("./a:chExt/@cx")
    assert clone_xfrm.xpath("./a:chExt/@cy") == src_xfrm.xpath("./a:chExt/@cy")
