"""Gherkin step implementations for font color features."""

from __future__ import annotations

from behave import given, then, when
from helpers import test_pptx

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR

font_color_pptx_path = test_pptx("font-color")


# given ===================================================


@given("a font with {color_type} color")
def step_given_font_with_color_type(context, color_type):
    context.textbox_idx = {"no": 0, "an RGB": 1, "a theme": 2}[color_type]
    context.prs = Presentation(font_color_pptx_path)
    textbox = context.prs.slides[0].shapes[context.textbox_idx]
    context.font = textbox.text_frame.paragraphs[0].runs[0].font


@given("a font with a color brightness setting of {setting}")
def step_font_with_color_brightness(context, setting):
    textbox_idx = {"no brightness adjustment": 2, "25% darker": 3, "40% lighter": 4}[setting]
    context.prs = Presentation(font_color_pptx_path)
    textbox = context.prs.slides[0].shapes[textbox_idx]
    context.font = textbox.text_frame.paragraphs[0].runs[0].font


# when ====================================================


@when("I set the font color brightness to {value}")
def step_set_font_color_brightness(context, value):
    context.font.color.brightness = float(value)


@when("I set the font {color_type} value")
def step_set_font_color_value(context, color_type):
    if color_type == "RGB":
        context.font.color.rgb = RGBColor(0x12, 0x34, 0x56)
    elif color_type == "theme color":
        context.font.color.theme_color = MSO_THEME_COLOR.DARK_1


# then ====================================================


@then("its color value matches its RGB color")
def step_color_value_matches_RGB_color(context):
    assert context.font.color.rgb == RGBColor(255, 102, 0)


@then("its color value matches its theme color")
def step_color_value_matches_theme_color(context):
    assert context.font.color.theme_color == MSO_THEME_COLOR.ACCENT_1


@then("the font's color type is {color_type}")
def step_then_font_color_type_is_value(context, color_type):
    expected_value = {
        "None": None,
        "RGB": MSO_COLOR_TYPE.RGB,
        "theme color": MSO_COLOR_TYPE.SCHEME,
    }[color_type]
    textbox = context.prs.slides[0].shapes[context.textbox_idx]
    font = textbox.text_frame.paragraphs[0].runs[0].font
    assert font.color.type == expected_value


@then("its color brightness value is {value}")
def step_color_brightness_value_matches(context, value):
    assert context.font.color.brightness == float(value)


@then("the font's {color_type} value matches the value I set")
def step_color_type_value_matches(context, color_type):
    textbox = context.prs.slides[0].shapes[context.textbox_idx]
    font = textbox.text_frame.paragraphs[0].runs[0].font
    if color_type == "RGB":
        assert font.color.rgb == RGBColor(0x12, 0x34, 0x56)
    else:
        assert font.color.theme_color == MSO_THEME_COLOR.DARK_1


@then("the font's color brightness is {value}")
def step_color_brightness_matches(context, value):
    textbox = context.prs.slides[0].shapes[context.textbox_idx]
    font = textbox.text_frame.paragraphs[0].runs[0].font
    assert font.color.brightness == float(value)


# -- effective font color (issue #938) ------------------------------------


@then("font.effective_color is {expected}")
def step_then_effective_color_is(context, expected):
    # -- evaluate the expected literal against the same symbols --
    expected_rgb = eval(expected, {"RGBColor": RGBColor})
    # -- re-resolve font from shape because `Given` step may have consumed it --
    textbox = context.prs.slides[0].shapes[context.textbox_idx]
    font = textbox.text_frame.paragraphs[0].runs[0].font
    rgb = font.effective_color
    assert rgb == expected_rgb, "expected %s, got %s" % (repr(expected_rgb), repr(rgb))


# -- hyperlink color override (issue #940) --------------------------------


@given("a font on a hyperlinked run with an explicit RGB color")
def step_given_hyperlinked_run_with_rgb(context):
    import io

    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = "click here"
    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    run.hyperlink.address = "https://example.com"

    # -- round-trip through a BytesIO so the marker must survive loading --
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    context.prs = Presentation(buf)
    context.font = (
        context.prs.slides[0].shapes[-1].text_frame.paragraphs[0].runs[0].font
    )


@given("a font on a hyperlinked run with theme color opted out")
def step_given_hyperlinked_run_with_opt_out(context):
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = "click here"
    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    run.hyperlink.address = "https://example.com"
    run.font.use_theme_hyperlink_color = False
    context.prs = prs
    context.font = run.font


@when("I assign {value} to font.use_theme_hyperlink_color")
def step_assign_use_theme_hyperlink_color(context, value):
    rhs = {"True": True, "False": False, "None": None}[value]
    context.font.use_theme_hyperlink_color = rhs


@then("font.use_theme_hyperlink_color is {value}")
def step_then_use_theme_hyperlink_color_is(context, value):
    expected = {"True": True, "False": False, "None": None}[value]
    run = context.prs.slides[0].shapes[-1].text_frame.paragraphs[0].runs[0]
    assert run.font.use_theme_hyperlink_color is expected, (
        "expected %r, got %r" % (expected, run.font.use_theme_hyperlink_color)
    )


@then("the run's explicit RGB color is preserved")
def step_then_rgb_preserved(context):
    run = context.prs.slides[0].shapes[-1].text_frame.paragraphs[0].runs[0]
    assert run.font.color.rgb == RGBColor(0xFF, 0x00, 0x00)
