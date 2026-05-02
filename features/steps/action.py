"""Gherkin step implementations for click action-related features."""

from __future__ import annotations

from behave import given, then, when
from helpers import test_file

from pptx import Presentation
from pptx.action import ActionSetting, Hyperlink, Sound
from pptx.enum.action import PP_ACTION

WAV_FILENAME = "act-click-sound.wav"

# given ===================================================


@given("an ActionSetting object having action {action} as click_action")
def given_an_ActionSetting_object_as_click_action(context, action):
    shape_idx = {"NONE": 0, "NAMED_SLIDE": 6}[action]
    slides = Presentation(test_file("act-props.pptm")).slides
    context.slides = slides
    context.click_action = slides[2].shapes[shape_idx].click_action


@given("another slide in the deck as slide")
def given_another_slide_in_the_deck_as_slide(context):
    context.slide = context.slides[1]


@given("a WAV sound attached as click_action.sound")
def given_a_WAV_sound_attached_as_click_action_sound(context):
    context.click_action.set_sound(test_file(WAV_FILENAME))


@given("a shape having click action {action}")
def given_a_shape_having_click_action_action(context, action):
    shape_idx = (
        "none",
        "first slide",
        "last slide",
        "previous slide",
        "next slide",
        "last slide viewed",
        "named slide",
        "end show",
        "hyperlink",
        "other presentation",
        "open file",
        "custom slide show",
        "OLE action",
        "run macro",
        "run program",
        "play media",
    ).index(action)
    slides = Presentation(test_file("act-props.pptm")).slides
    context.slides = slides
    shape = slides[2].shapes[shape_idx]
    context.shape = shape
    context.click_action = shape.click_action
    context.hover_action = shape.hover_action


# when ====================================================


@when("I assign {value} to click_action.hyperlink.address")
def when_I_assign_value_to_click_action_hyperlink_address(context, value):
    value = None if value == "None" else value
    context.click_action.hyperlink.address = value


@when("I assign {value} to click_action.target_slide")
def when_I_assign_value_to_click_action_target_slide(context, value):
    rhs = {"None": None, "slide": context.slide}[value]
    context.click_action.target_slide = rhs


@when("I call click_action.set_sound with a WAV file")
def when_I_call_click_action_set_sound_with_a_WAV_file(context):
    context.click_action.set_sound(test_file(WAV_FILENAME))


@when("I call click_action.remove_sound")
def when_I_call_click_action_remove_sound(context):
    context.click_action.remove_sound()


@when("I assign {value} to hover_action.hyperlink.address")
def when_I_assign_value_to_hover_action_hyperlink_address(context, value):
    value = None if value == "None" else value
    context.hover_action.hyperlink.address = value


@when("I assign {value} to hover_action.target_slide")
def when_I_assign_value_to_hover_action_target_slide(context, value):
    rhs = {"None": None, "slide": context.slide}[value]
    context.hover_action.target_slide = rhs


@when("I assign {value} to hover_action.screen_tip")
def when_I_assign_value_to_hover_action_screen_tip(context, value):
    context.hover_action.screen_tip = None if value == "None" else value


# then ====================================================


@then("click_action.action is {member_name}")
def then_click_action_action_is_value(context, member_name):
    click_action = context.click_action
    expected_value = getattr(PP_ACTION, member_name)
    assert click_action.action == expected_value


@then("click_action.hyperlink is a Hyperlink object")
def then_click_action_hyperlink_is_a_Hyperlink_object(context):
    hyperlink = context.click_action.hyperlink
    assert isinstance(hyperlink, Hyperlink)


@then("click_action.hyperlink.address is {value}")
def then_click_action_hyperlink_address_is_value(context, value):
    expected_value = None if value == "None" else value
    hyperlink = context.click_action.hyperlink
    assert hyperlink.address == expected_value, "expected %s, got %s" % (
        expected_value,
        hyperlink.address,
    )


@then("click_action.sound is a Sound object")
def then_click_action_sound_is_a_Sound_object(context):
    assert isinstance(context.click_action.sound, Sound)


@then("click_action.sound is None")
def then_click_action_sound_is_None(context):
    assert context.click_action.sound is None


@then("click_action.sound.name is the WAV filename")
def then_click_action_sound_name_is_the_WAV_filename(context):
    assert context.click_action.sound.name == WAV_FILENAME


@then("click_action.sound.blob is the WAV bytes")
def then_click_action_sound_blob_is_the_WAV_bytes(context):
    with open(test_file(WAV_FILENAME), "rb") as f:
        expected = f.read()
    assert context.click_action.sound.blob == expected


@then("click_action.target_slide is {value}")
def then_click_action_target_slide_is_value(context, value):
    if value.startswith("slides["):
        idx = value[7]
        expected_value = context.slides[int(idx)]
    elif value == "None":
        expected_value = None
    else:
        expected_value = context.slide

    click_action = context.click_action
    assert click_action.target_slide == expected_value


@then("hover_action is an ActionSetting object")
def then_hover_action_is_an_ActionSetting_object(context):
    assert isinstance(context.hover_action, ActionSetting)


@then("hover_action.action is {member_name}")
def then_hover_action_action_is_value(context, member_name):
    expected_value = getattr(PP_ACTION, member_name)
    assert context.hover_action.action == expected_value, (
        "expected %s, got %s" % (expected_value, context.hover_action.action)
    )


@then("hover_action.hyperlink.address is {value}")
def then_hover_action_hyperlink_address_is_value(context, value):
    expected_value = None if value == "None" else value
    address = context.hover_action.hyperlink.address
    assert address == expected_value, (
        "expected %s, got %s" % (expected_value, address)
    )


@then("hover_action.target_slide is {value}")
def then_hover_action_target_slide_is_value(context, value):
    if value.startswith("slides["):
        idx = value[7]
        expected_value = context.slides[int(idx)]
    elif value == "None":
        expected_value = None
    else:
        expected_value = context.slide

    assert context.hover_action.target_slide == expected_value


@then("hover_action.screen_tip is {value}")
def then_hover_action_screen_tip_is_value(context, value):
    expected_value = None if value == "None" else value
    assert context.hover_action.screen_tip == expected_value, (
        "expected %r, got %r" % (expected_value, context.hover_action.screen_tip)
    )
