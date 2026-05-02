"""Gherkin step implementations for ShadowFormat-related features."""

from __future__ import annotations

from behave import given, then, when
from helpers import test_pptx

from pptx import Presentation
from pptx.dml.effect import EffectFormat, GlowFormat, ReflectionFormat, SoftEdgeFormat

# given ====================================================


@given("a ShadowFormat object that {inherits} as shadow")
def given_a_ShadowFormat_object_that_inherits_or_not(context, inherits):
    shape_idx = {"inherits": 0, "does not inherit": 1}[inherits]
    shape = Presentation(test_pptx("dml-effect")).slides[0].shapes[shape_idx]
    context.shape = shape
    context.shadow = shape.shadow


# when =====================================================


@when("I assign {value} to shadow.inherit")
def when_I_assign_value_to_shadow_inherit(context, value):
    context.shadow.inherit = eval(value)


@when("I assign {value} to shadow.blur_radius")
def when_I_assign_value_to_shadow_blur_radius(context, value):
    context.shadow.blur_radius = int(value)


# then =====================================================


@then("shadow.inherit is {bool_str}")
def then_shadow_inherit_is_bool_val(context, bool_str):
    expected_value = eval(bool_str)
    actual_value = context.shadow.inherit
    assert actual_value is expected_value, "shadow.inherit is %s" % actual_value


@then("shadow.blur_radius is {value}")
def then_shadow_blur_radius_is(context, value):
    expected_value = int(value)
    actual_value = context.shadow.blur_radius
    assert actual_value == expected_value, "shadow.blur_radius is %s" % actual_value


@then("the {effect} effect on that shape inherits")
def then_the_effect_on_that_shape_inherits(context, effect):
    effects = EffectFormat(context.shape._element.spPr)
    sub = getattr(effects, effect)
    assert isinstance(sub, (GlowFormat, ReflectionFormat, SoftEdgeFormat)), (
        "unexpected effect type %s" % type(sub).__name__
    )
    assert sub.inherit is True, "%s effect does not inherit" % effect


@then('the sibling a:effectRef has idx "{idx}"')
def then_the_sibling_a_effectRef_has_idx(context, idx):
    sp = context.shape._element
    effectRefs = sp.xpath("./p:style/a:effectRef")
    assert effectRefs, "no sibling p:style/a:effectRef found on the shape"
    actual = effectRefs[0].get("idx")
    assert actual == idx, 'a:effectRef/@idx is "%s", expected "%s"' % (actual, idx)
