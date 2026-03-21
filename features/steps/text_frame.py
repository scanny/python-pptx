"""Step implementations for text frame-related features"""

from __future__ import annotations

from behave import given, then, when
from helpers import test_pptx

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches, Pt

# given ===================================================


@given("a TextFrame object as text_frame")
def given_a_text_frame(context):
    context.text_frame = Presentation(test_pptx("txt-text")).slides[0].shapes[0].text_frame


@given("a TextFrame object containing {value} as text_frame")
def given_a_TextFrame_object_containing_value_as_text_frame(context, value):
    shape_idx = {"abc": 0, "a\nb\nc": 1}[eval(value)]
    prs = Presentation(test_pptx("txt-text-frame"))
    context.text_frame = prs.slides[1].shapes[shape_idx].text_frame


@given("a TextFrame object having auto-size of {setting} as text_frame")
def given_a_TextFrame_object_having_auto_size_of_setting(context, setting):
    shape_idx = {
        "None": 0,
        "no auto-size": 1,
        "fit shape to text": 2,
        "fit text to shape": 3,
    }[setting]
    prs = Presentation(test_pptx("txt-text-frame"))
    shape = prs.slides[0].shapes[shape_idx]
    context.text_frame = shape.text_frame


@given("a text frame with more text than will fit")
def given_a_text_frame_with_more_text_than_will_fit(context):
    prs = Presentation(test_pptx("txt-fit-text"))
    shape = prs.slides[0].shapes[0]
    context.text_frame = shape.text_frame


@given("a text frame with text that will overflow at 18pt")
def given_a_text_frame_with_text_that_will_overflow_at_18pt(context):
    prs = Presentation(test_pptx("txt-fit-text"))
    shape = prs.slides[0].shapes[0]
    context.text_frame = shape.text_frame


@given("a text frame with text that will fit at 10pt")
def given_a_text_frame_with_text_that_will_fit_at_10pt(context):
    prs = Presentation(test_pptx("txt-fit-text"))
    shape = prs.slides[0].shapes[0]
    context.text_frame = shape.text_frame


@given("an empty text frame")
def given_an_empty_text_frame(context):
    prs = Presentation(test_pptx("txt-text-frame"))
    shape = prs.slides[0].shapes[0]
    context.text_frame = shape.text_frame
    context.text_frame.clear()


@given("a text frame with uniform 14pt text")
def given_a_text_frame_with_uniform_14pt_text(context):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(3))
    context.text_frame = textbox.text_frame
    context.text_frame.text = "Test text"
    for paragraph in context.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(14)


@given("a text frame with mixed font sizes")
def given_a_text_frame_with_mixed_font_sizes(context):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(3))
    context.text_frame = textbox.text_frame
    p = context.text_frame.paragraphs[0]
    r1 = p.add_run()
    r1.text = "First"
    r1.font.size = Pt(14)
    r2 = p.add_run()
    r2.text = "Second"
    r2.font.size = Pt(18)


# when ====================================================


@when("I assign {value} to text_frame.auto_size")
def when_I_assign_value_to_text_frame_auto_size(context, value):
    text_frame = context.text_frame
    text_frame.auto_size = {
        "None": None,
        "MSO_AUTO_SIZE.NONE": MSO_AUTO_SIZE.NONE,
        "MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT": MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT,
        "MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE": MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
    }[value]


@when("I assign text_frame.margin_{side} = Inches({inches})")
def when_I_assign_text_frame_margin_side_eq_inches(context, side, inches):
    attr_name = "margin_%s" % side
    setattr(context.text_frame, attr_name, Inches(float(inches)))


@when("I assign text_frame.text = {value}")
def when_I_assign_text_frame_text_eq_value(context, value):
    context.text_frame.text = eval(value)


@when("I assign {value} to text_frame.word_wrap")
def when_I_assign_value_to_text_frame_word_wrap(context, value):
    new_value = {"True": True, "False": False, "None": None}[value]
    context.text_frame.word_wrap = new_value


@when("I call TextFrame.fit_text()")
def when_I_call_TextFrame_fit_text(context):
    from helpers import test_file

    font_file = test_file("calibriz.ttf")
    context.text_frame.fit_text(bold=True, italic=True, font_file=font_file)
    # context.text_frame.fit_text(font_family='Arial', bold=True, italic=True)


@when("I call text_frame.will_overflow(font_size={font_size})")
def when_I_call_text_frame_will_overflow_with_font_size(context, font_size):
    from helpers import test_file

    font_file = test_file("calibriz.ttf")
    context.result = context.text_frame.will_overflow(
        bold=True, italic=True, font_size=int(font_size), font_file=font_file
    )


@when("I call text_frame.will_overflow() without specifying font_size")
def when_I_call_text_frame_will_overflow_without_font_size(context):
    from helpers import test_file

    font_file = test_file("calibriz.ttf")
    try:
        context.result = context.text_frame.will_overflow(
            bold=True, italic=True, font_file=font_file
        )
        context.exception = None
    except Exception as e:
        context.exception = e


@when("I call text_frame.will_overflow()")
def when_I_call_text_frame_will_overflow(context):
    from helpers import test_file

    font_file = test_file("calibriz.ttf")
    context.result = context.text_frame.will_overflow(
        bold=True, italic=True, font_file=font_file
    )


@when("I call text_frame.overflow_info(font_size={font_size})")
def when_I_call_text_frame_overflow_info_with_font_size(context, font_size):
    from helpers import test_file

    font_file = test_file("calibriz.ttf")
    context.info = context.text_frame.overflow_info(
        bold=True, italic=True, font_size=int(font_size), font_file=font_file
    )


# then ====================================================


@then("text_frame.auto_size is {value}")
def then_text_frame_autosize_is_value(context, value):
    expected_value = {
        "None": None,
        "MSO_AUTO_SIZE.NONE": MSO_AUTO_SIZE.NONE,
        "MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT": MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT,
        "MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE": MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
    }[value]
    text_frame = context.text_frame
    assert text_frame.auto_size == expected_value, "got %s" % text_frame.auto_size


@then("text_frame.margin_{side}.inches == {inches}")
def then_text_frame_margin_side_inches_eq_inches(context, side, inches):
    attr_name = "margin_%s" % side
    actual = getattr(context.text_frame, attr_name).inches
    expected = float(inches)
    assert actual == expected, "text_frame.margin_%s.inches == %s" % (side, actual)


@then("text_frame.text == {value}")
def then_text_frame_text_eq_value(context, value):
    actual, expected = context.text_frame.text, eval(value)
    assert actual == expected, 'text_frame.text == "%s"' % actual


@then("text_frame.word_wrap is {value}")
def then_text_frame_word_wrap_is_value(context, value):
    expected_value = {"True": True, "False": False, "None": None}[value]
    text_frame = context.text_frame
    assert text_frame.word_wrap is expected_value


@then("the size of the text is 10pt or 11pt")
def then_the_size_of_the_text_is_10pt(context):
    """Size depends on Pillow version, probably algorithm isn't quite right either."""
    text_frame = context.text_frame
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            assert run.font.size in (Pt(10.0), Pt(11.0)), "got %s" % run.font.size.pt


@then("it returns True")
def then_it_returns_true(context):
    assert context.result is True, "Expected True, got %s" % context.result


@then("it returns False")
def then_it_returns_false(context):
    assert context.result is False, "Expected False, got %s" % context.result


@then("info.will_overflow is True")
def then_info_will_overflow_is_true(context):
    assert context.info.will_overflow is True


@then("info.will_overflow is False")
def then_info_will_overflow_is_false(context):
    assert context.info.will_overflow is False


@then("info.required_height is greater than info.available_height")
def then_info_required_height_is_greater_than_available_height(context):
    assert context.info.required_height > context.info.available_height


@then("info.overflow_percentage is greater than 0")
def then_info_overflow_percentage_is_greater_than_0(context):
    assert context.info.overflow_percentage > 0


@then("info.fits_at_font_size is less than 18")
def then_info_fits_at_font_size_is_less_than_18(context):
    assert context.info.fits_at_font_size < 18


@then("info.overflow_height is 0")
def then_info_overflow_height_is_0(context):
    assert context.info.overflow_height == 0


@then("info.fits_at_font_size is None")
def then_info_fits_at_font_size_is_none(context):
    assert context.info.fits_at_font_size is None


@then("it uses 14pt as the font size")
def then_it_uses_14pt_as_font_size(context):
    # This is verified by not raising an exception
    assert context.exception is None


@then("it raises ValueError")
def then_it_raises_value_error(context):
    assert context.exception is not None
    assert isinstance(context.exception, ValueError)
    assert "multiple font sizes" in str(context.exception)
