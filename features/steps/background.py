"""Gherkin step implementations for slide background-related features."""

from __future__ import annotations

from behave import given, then, when
from helpers import test_pptx
from lxml import etree

from pptx import Presentation
from pptx.oxml.ns import qn

# given ===================================================


@given("a _Background object having {type} background as background")
def given_a_Background_object_having_type_background(context, type):
    sld_idx = {"no": 0, "a fill": 1, "a style reference": None}[type]
    prs = Presentation(test_pptx("sld-background"))
    # -- retain the parent presentation so the slide objects stay alive --
    context.prs = prs
    slide = prs.slide_masters[0] if sld_idx is None else prs.slides[sld_idx]
    context.slide = slide
    context.background = slide.background


@given("a Slide object with no explicit background")
def given_a_Slide_object_with_no_explicit_background(context):
    prs = Presentation(test_pptx("sld-background"))
    context.dst_prs = prs
    context.dst_slide = prs.slides[0]  # -- fixture slide 0 has no explicit bg --
    assert context.dst_slide.background.bg_element is None


@given("a Slide object with an explicit fill background")
def given_a_Slide_object_with_an_explicit_fill_background(context):
    prs = Presentation(test_pptx("sld-background"))
    context.dst_prs = prs
    context.dst_slide = prs.slides[1]  # -- fixture slide 1 has an explicit fill --
    assert context.dst_slide.background.bg_element is not None


@given("a source Slide object with an explicit fill background")
def given_a_source_Slide_with_an_explicit_fill_background(context):
    prs = Presentation(test_pptx("sld-background"))
    context.src_prs = prs
    context.src_slide = prs.slides[1]
    assert context.src_slide.background.bg_element is not None


@given("a source Slide object with no explicit background")
def given_a_source_Slide_with_no_explicit_background(context):
    prs = Presentation(test_pptx("sld-background"))
    context.src_prs = prs
    context.src_slide = prs.slides[0]
    assert context.src_slide.background.bg_element is None


# when ====================================================


@when("I call slide.copy_background_from(source)")
def when_I_call_slide_copy_background_from(context):
    context.dst_slide.copy_background_from(context.src_slide)


# then ====================================================


@then("background.fill is a FillFormat object")
def then_background_fill_is_a_Fill_object(context):
    cls_name = context.background.fill.__class__.__name__
    assert cls_name == "FillFormat", "background.fill is a %s object" % cls_name


@then("background.bg_element is the p:bg element")
def then_background_bg_element_is_p_bg(context):
    bg = context.background.bg_element
    assert bg is not None, "background.bg_element is None"
    assert bg.tag == qn("p:bg"), "expected p:bg, got %s" % bg.tag


@then("background.bg_element is None")
def then_background_bg_element_is_None(context):
    assert context.background.bg_element is None, (
        "expected None, got %r" % context.background.bg_element
    )


@then("the destination slide has an explicit p:bg element")
def then_destination_has_explicit_bg(context):
    bg = context.dst_slide.background.bg_element
    assert bg is not None, "destination slide has no explicit p:bg"


@then("the destination slide has no explicit p:bg element")
def then_destination_has_no_explicit_bg(context):
    bg = context.dst_slide.background.bg_element
    assert bg is None, "destination slide still has a p:bg element"


@then("the destination p:bg subtree matches the source")
def then_destination_bg_matches_source(context):
    dst_bg = context.dst_slide.background.bg_element
    src_bg = context.src_slide.background.bg_element
    assert dst_bg is not None
    assert src_bg is not None
    dst_c14n = etree.tostring(dst_bg, method="c14n")
    src_c14n = etree.tostring(src_bg, method="c14n")
    assert dst_c14n == src_c14n, (
        "destination p:bg differs from source p:bg\nDST: %s\nSRC: %s"
        % (dst_c14n, src_c14n)
    )
