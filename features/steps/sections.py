"""Gherkin step implementations for presentation-section features."""

from __future__ import annotations

from behave import given, then, when
from behave.runner import Context
from helpers import test_pptx

from pptx import Presentation
from pptx.oxml.ns import qn

# given ===================================================


@given("a presentation with three slides")
def given_a_presentation_with_three_slides(context: Context):
    context.prs = Presentation(test_pptx("sld-slides"))


# when ====================================================


@when('I add a section named "{name}" with slides 1 and 2')
def when_add_section_with_two_slides(context: Context, name: str):
    slides = [context.prs.slides[0], context.prs.slides[1]]
    context.prs.sections.add_section(name, slides=slides)


@when('I rename the first section to "{new_name}"')
def when_rename_first_section(context: Context, new_name: str):
    context.prs.sections[0].name = new_name


@when("I remove the first section")
def when_remove_first_section(context: Context):
    section = context.prs.sections[0]
    context.prs.sections.remove(section)


# then ====================================================


@then("Presentation.sections has {count:d} section")
@then("Presentation.sections has {count:d} sections")
def then_sections_has_count(context: Context, count: int):
    actual = len(context.prs.sections)
    assert actual == count, "expected %d sections, got %d" % (count, actual)


@then('the section is named "{name}"')
def then_section_is_named(context: Context, name: str):
    actual = context.prs.sections[0].name
    assert actual == name, "expected %r, got %r" % (name, actual)


@then("the section has two assigned slides")
def then_section_has_two_assigned_slides(context: Context):
    section = context.prs.sections[0]
    slides = section.slides
    assert len(slides) == 2, "expected 2 assigned slides, got %d" % len(slides)
    # -- they should be the first two slides of the deck --
    assert slides[0] == context.prs.slides[0]
    assert slides[1] == context.prs.slides[1]


@then("the presentation.xml has a p14:sectionLst under p:extLst")
def then_prs_xml_has_sectionLst(context: Context):
    prs_element = context.prs._element  # pyright: ignore[reportPrivateUsage]
    extLst = prs_element.find(qn("p:extLst"))
    assert extLst is not None, "no p:extLst present on presentation"
    sectionLst = extLst.find("./{0}/{1}".format(qn("p:ext"), qn("p14:sectionLst")))
    assert sectionLst is not None, "no p14:sectionLst under p:extLst/p:ext"


@then("the presentation has no p:extLst element")
def then_prs_has_no_extLst(context: Context):
    prs_element = context.prs._element  # pyright: ignore[reportPrivateUsage]
    assert prs_element.find(qn("p:extLst")) is None
