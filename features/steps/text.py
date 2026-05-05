"""Gherkin step implementations for text-related features."""

from __future__ import annotations

from behave import given, then, when
from helpers import test_pptx

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, PP_AUTO_NUMBER
from pptx.util import Emu, Pt

# given ===================================================


@given("a _Paragraph object as paragraph")
def given_a_Paragraph_object_as_paragraph(context):
    prs = Presentation(test_pptx("txt-text"))
    context.paragraph = prs.slides[0].shapes[0].text_frame.paragraphs[0]


@given("a _Paragraph object containing {value} as paragraph")
def given_a_Paragraph_object_containing_value_as_paragraph(context, value):
    prs = Presentation(test_pptx("txt-text"))
    paragraph_idx = {"abc": 0, "a\vb\vc": 1}[eval(value)]
    context.paragraph = prs.slides[0].shapes[1].text_frame.paragraphs[paragraph_idx]


@given("a _Paragraph with a math equation bracketed by plain-text runs as paragraph")
def given_a_Paragraph_with_a_math_equation_bracketed_by_plain_text_runs(context):
    # -- authors an `a:p` that matches the shape PowerPoint emits when a user types
    # -- math syntax like ``m^3`` and PowerPoint auto-formats it: an
    # -- ``mc:AlternateContent`` wrapper with an ``a14:m/m:oMath`` live side and an
    # -- ``mc:Fallback`` carrying the plain-text rendering. Previously
    # -- ``_Paragraph.text`` dropped the fallback text; the fix includes it. See
    # -- issue #947.
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    tb = slide.shapes.add_textbox(914400, 914400, 4000000, 1000000)
    paragraph = tb.text_frame.paragraphs[0]
    paragraph.add_run().text = "Formula: "
    paragraph.add_math_equation(
        "<m:oMathPara xmlns:m="
        '"http://schemas.openxmlformats.org/officeDocument/2006/math">'
        "<m:oMath>"
        "<m:r><m:t>E=mc</m:t></m:r>"
        "<m:sSup><m:e><m:r><m:t></m:t></m:r></m:e>"
        "<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>"
        "</m:oMath>"
        "</m:oMathPara>"
    )
    paragraph.add_run().text = " done."
    context.paragraph = paragraph


@given("a paragraph having line spacing of {setting}")
def given_a_paragraph_having_line_spacing_of_setting(context, setting):
    paragraph_idx = {"no explicit setting": 0, "1.5 lines": 1, "20 pt": 2}[setting]
    prs = Presentation(test_pptx("txt-paragraph-spacing"))
    text_frame = prs.slides[2].shapes[0].text_frame
    context.paragraph = text_frame.paragraphs[paragraph_idx]


@given("a paragraph having space {before_after} of {setting}")
def given_a_paragraph_having_space_before_after_of_setting(context, before_after, setting):
    slide_idx = {"before": 0, "after": 1}[before_after]
    paragraph_idx = {"no explicit setting": 0, "6 pt": 1}[setting]
    prs = Presentation(test_pptx("txt-paragraph-spacing"))
    text_frame = prs.slides[slide_idx].shapes[0].text_frame
    context.paragraph = text_frame.paragraphs[paragraph_idx]


@given("a _Run object as run")
def given_a_Run_object_as_run(context):
    prs = Presentation(test_pptx("txt-text"))
    context.run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]


@given("a _Run object containing text as run")
def given_a_Run_object_containing_text_as_run(context):
    prs = Presentation(test_pptx("txt-text"))
    context.run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]


@given("a text run")
def given_a_text_run(context):
    prs = Presentation(test_pptx("txt-text"))
    context.prs = prs
    context.run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]


@given("a text run in a table cell")
def given_a_text_run_in_a_table_cell(context):
    prs = Presentation(test_pptx("txt-text"))
    cell = prs.slides[1].shapes[0].table.cell(0, 0)
    context.run = cell.text_frame.paragraphs[0].runs[0]


@given("a text run having a hyperlink")
def given_a_text_run_having_a_hyperlink(context):
    prs = Presentation(test_pptx("txt-text"))
    context.prs = prs
    context.run = prs.slides[0].shapes[0].text_frame.paragraphs[1].runs[0]


@given("another slide in the deck as target_slide")
def given_another_slide_in_the_deck_as_target_slide(context):
    # -- reuse the presentation loaded by "a text run" / "a text run having a
    # -- hyperlink". Use slide[1] as the hop target. --
    context.target_slide = context.prs.slides[1]


@given("a text run with a slide-jump hyperlink to another slide")
def given_a_text_run_with_a_slide_jump_hyperlink(context):
    prs = Presentation(test_pptx("txt-text"))
    context.prs = prs
    context.target_slide = prs.slides[1]
    context.run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    context.run.hyperlink.target_slide = context.target_slide


# when ====================================================


@when("I assign a typeface name to the font")
def when_assign_typeface_name_to_font(context):
    context.font.name = "Verdana"


@when("I assign None to hyperlink.address")
def when_assign_None_to_hyperlink_address(context):
    context.run.hyperlink.address = None


@when("I assign paragraph.alignment = PP_ALIGN.CENTER")
def when_I_assign_paragraph_alignment_eq_center(context):
    context.paragraph.alignment = PP_ALIGN.CENTER


@when("I assign paragraph.level = 1")
def when_I_assign_paragraph_leve_eq_1(context):
    context.paragraph.level = 1


@when("I assign paragraph.text = {value}")
def when_I_assign_paragraph_text_eq_value(context, value):
    context.paragraph.text = eval(value)


@when("I assign run.text = {value}")
def when_I_assign_run_text_eq_value(context, value):
    context.run.text = eval(value)


@when("I assign {value_str} to paragraph.line_spacing")
def when_I_assign_value_to_paragraph_line_spacing(context, value_str):
    value = {
        "1.5": 1.5,
        "2.0": 2.0,
        "254000": Emu(254000),
        "304800": Emu(304800),
        "None": None,
    }[value_str]
    paragraph = context.paragraph
    paragraph.line_spacing = value


@when("I assign {value_str} to paragraph.space_{before_after}")
def when_I_assign_value_to_paragraph_space_before_after(context, value_str, before_after):
    value = {"76200": 76200, "38100": 38100, "None": None}[value_str]
    attr_name = {"before": "space_before", "after": "space_after"}[before_after]
    paragraph = context.paragraph
    setattr(paragraph, attr_name, value)


@when("I set the hyperlink address")
def when_set_hyperlink_address(context):
    context.run_text = "python-pptx @ GitHub"
    context.address = "https://github.com/scanny/python-pptx"

    run = context.run
    run.text = context.run_text
    hlink = run.hyperlink
    hlink.address = context.address


@when("I assign target_slide to run.hyperlink.target_slide")
def when_assign_target_slide_to_run_hyperlink_target_slide(context):
    context.run.hyperlink.target_slide = context.target_slide


@when("I assign None to run.hyperlink.target_slide")
def when_assign_None_to_run_hyperlink_target_slide(context):
    context.run.hyperlink.target_slide = None


# then ====================================================


@then("paragraph.alignment == PP_ALIGN.CENTER")
def then_paragraph_alignment_eq_center(context):
    actual, expected = context.paragraph.alignment, PP_ALIGN.CENTER
    assert actual == expected, "paragraph.alignment == %s" % actual


@then("paragraph.level == 1")
def then_paragraph_level_eq_1(context):
    actual, expected = context.paragraph.level, 1
    assert actual == expected, "paragraph.level == %s" % actual


@then("paragraph.line_spacing is {value_str}")
def then_paragraph_line_spacing_is_value(context, value_str):
    value = {
        "None": None,
        "1.0": 1.0,
        "1.5": 1.5,
        "2.0": 2.0,
        "254000": 254000,
        "304800": 304800,
    }[value_str]
    paragraph = context.paragraph
    assert paragraph.line_spacing == value


@then("paragraph.line_spacing.pt {result}")
def then_paragraph_line_spacing_pt_result(context, result):
    value, exception = {
        "raises AttributeError": (None, AttributeError),
        "is 20.0": (20.0, None),
        "is 24.0": (24.0, None),
    }[result]
    line_spacing = context.paragraph.line_spacing
    if value is not None:
        assert line_spacing.pt == value
    if exception is not None:
        try:
            line_spacing.pt
            raise AssertionError("did not raise")
        except exception:
            pass


@then("paragraph.space_{before_after} is {value_str}")
def then_paragraph_space_before_is_value(context, before_after, value_str):
    attr_name = {"before": "space_before", "after": "space_after"}[before_after]
    value = None if value_str == "None" else int(value_str)
    paragraph = context.paragraph
    assert getattr(paragraph, attr_name) == value


@then("paragraph.space_{before_after}.pt {result}")
def then_paragraph_space_before_pt_is_value(context, before_after, result):
    attr_name = {"before": "space_before", "after": "space_after"}[before_after]
    value, exception = {
        "raises AttributeError": (None, AttributeError),
        "== 6.0": (6.0, None),
    }[result]
    space_before_after = getattr(context.paragraph, attr_name)
    if value is not None:
        assert space_before_after.pt == value
    if exception is not None:
        try:
            space_before_after.pt
            raise AssertionError("did not raise")
        except exception:
            pass


@then("paragraph.text == {value}")
def then_paragraph_text_eq_value(context, value):
    actual, expected = context.paragraph.text, eval(value)
    assert actual == expected, 'paragraph.text == "%s"' % actual


@then("paragraph.text matches the assigned string")
def then_paragraph_text_matches_the_assigned_string(context):
    paragraph = context.paragraph
    assert paragraph.text == " Boo Far \n Faz Foo "


@then("run.text == {value}")
def then_run_text_is_value(context, value):
    actual, expected = context.run.text, eval(value)
    assert actual == expected, "run.text == %s" % (actual,)


@then("run.text is a hyperlink")
def then_run_text_is_a_hyperlink(context):
    run = context.run
    hlink = run.hyperlink
    assert run.text == context.run_text
    assert hlink.address == context.address


@then("run.text is not a hyperlink")
def then_run_text_is_not_a_hyperlink(context):
    hlink = context.run.hyperlink
    assert hlink.address is None


@then("run.hyperlink.target_slide is target_slide")
def then_run_hyperlink_target_slide_is_target_slide(context):
    actual = context.run.hyperlink.target_slide
    assert actual == context.target_slide, (
        "run.hyperlink.target_slide == %r (expected %r)"
        % (actual, context.target_slide)
    )


@then("run.hyperlink.target_slide is None")
def then_run_hyperlink_target_slide_is_None(context):
    actual = context.run.hyperlink.target_slide
    assert actual is None, "run.hyperlink.target_slide == %r" % (actual,)


@then("run.hyperlink.address is None")
def then_run_hyperlink_address_is_None(context):
    actual = context.run.hyperlink.address
    assert actual is None, "run.hyperlink.address == %r" % (actual,)


@then("the font name matches the typeface I set")
def then_font_name_matches_typeface_I_set(context):
    assert context.font.name == "Verdana"


# bullet steps ============================================


@when('I call paragraph.bullet.character("{char}")')
def when_I_call_paragraph_bullet_character(context, char):
    context.paragraph.bullet.character(char)


@when("I call paragraph.bullet.auto_number(PP_AUTO_NUMBER.{spec})")
def when_I_call_paragraph_bullet_auto_number(context, spec):
    # -- `spec` is either "MEMBER_NAME" or "MEMBER_NAME, <int>" --
    if "," in spec:
        member, start_at_str = (s.strip() for s in spec.split(","))
        start_at = int(start_at_str)
    else:
        member, start_at = spec.strip(), None
    scheme = getattr(PP_AUTO_NUMBER, member)
    context.paragraph.bullet.auto_number(scheme, start_at)


@when("I call paragraph.bullet.none()")
def when_I_call_paragraph_bullet_none(context):
    context.paragraph.bullet.none()


@when("I call paragraph.bullet.clear()")
def when_I_call_paragraph_bullet_clear(context):
    context.paragraph.bullet.clear()


@when('I assign paragraph.bullet.font = "{typeface}"')
def when_assign_paragraph_bullet_font(context, typeface):
    context.paragraph.bullet.font = typeface


@when("I assign paragraph.bullet.size_pct = {value:f}")
def when_assign_paragraph_bullet_size_pct(context, value):
    context.paragraph.bullet.size_pct = value


@when("I assign paragraph.bullet.size_points = Pt({value:d})")
def when_assign_paragraph_bullet_size_points(context, value):
    from pptx.util import Pt

    context.paragraph.bullet.size_points = Pt(value)


@when("I assign RGB FF0000 to paragraph.bullet.color")
def when_assign_rgb_ff0000_to_paragraph_bullet_color(context):
    from pptx.dml.color import RGBColor

    context.paragraph.bullet.color.rgb = RGBColor(0xFF, 0x00, 0x00)


@then("paragraph.bullet.type is None")
def then_paragraph_bullet_type_is_None(context):
    assert context.paragraph.bullet.type is None, (
        "paragraph.bullet.type == %r" % context.paragraph.bullet.type
    )


@then('paragraph.bullet.type == "{expected}"')
def then_paragraph_bullet_type_eq(context, expected):
    actual = context.paragraph.bullet.type
    assert actual == expected, "paragraph.bullet.type == %r" % actual


@then('paragraph.bullet.char == "{expected}"')
def then_paragraph_bullet_char_eq(context, expected):
    actual = context.paragraph.bullet.char
    assert actual == expected, "paragraph.bullet.char == %r" % actual


@then("paragraph.bullet.char is None")
def then_paragraph_bullet_char_is_None(context):
    assert context.paragraph.bullet.char is None, (
        "paragraph.bullet.char == %r" % context.paragraph.bullet.char
    )


@then("paragraph.bullet.number_scheme == PP_AUTO_NUMBER.{member}")
def then_paragraph_bullet_number_scheme_eq(context, member):
    expected = getattr(PP_AUTO_NUMBER, member)
    actual = context.paragraph.bullet.number_scheme
    assert actual == expected, "paragraph.bullet.number_scheme == %r" % actual


@then("paragraph.bullet.start_at == {expected:d}")
def then_paragraph_bullet_start_at_eq(context, expected):
    actual = context.paragraph.bullet.start_at
    assert actual == expected, "paragraph.bullet.start_at == %r" % actual


@then('paragraph.bullet.font == "{expected}"')
def then_paragraph_bullet_font_eq(context, expected):
    actual = context.paragraph.bullet.font
    assert actual == expected, "paragraph.bullet.font == %r" % actual


@then("paragraph.bullet.size_pct == {expected:f}")
def then_paragraph_bullet_size_pct_eq(context, expected):
    actual = context.paragraph.bullet.size_pct
    assert actual == expected, "paragraph.bullet.size_pct == %r" % actual


@then("paragraph.bullet.size_points.pt == {expected:f}")
def then_paragraph_bullet_size_points_pt_eq(context, expected):
    actual = context.paragraph.bullet.size_points.pt
    assert actual == expected, "paragraph.bullet.size_points.pt == %r" % actual


@then("paragraph.bullet.color.rgb == RGBColor(FF0000)")
def then_paragraph_bullet_color_rgb_eq_ff0000(context):
    from pptx.dml.color import RGBColor

    actual = context.paragraph.bullet.color.rgb
    assert actual == RGBColor(0xFF, 0x00, 0x00), (
        "paragraph.bullet.color.rgb == %r" % actual
    )


# -- issue #134: one-call add_run / add_paragraph ----------------------


@when("I call paragraph.add_run with text and font kwargs")
def when_I_call_paragraph_add_run_with_kwargs(context):
    context.run = context.paragraph.add_run(
        "bold-24pt-red",
        bold=True,
        size=Pt(24),
        color=RGBColor(0xFF, 0x00, 0x00),
    )


@then("the new run carries the text and font kwargs")
def then_new_run_carries_text_and_font_kwargs(context):
    run = context.run
    assert run.text == "bold-24pt-red", "run.text == %r" % run.text
    assert run.font.bold is True, "run.font.bold is %r" % run.font.bold
    assert run.font.size == Pt(24), "run.font.size == %r" % run.font.size
    assert run.font.color.rgb == RGBColor(0xFF, 0x00, 0x00), (
        "run.font.color.rgb == %r" % run.font.color.rgb
    )


@when("I call text_frame.add_paragraph with text and font kwargs")
def when_I_call_text_frame_add_paragraph_with_kwargs(context):
    context.paragraph = context.text_frame.add_paragraph(
        "styled line",
        italic=True,
        font_name="Arial",
    )


@then("the new paragraph's run carries the text and font kwargs")
def then_new_paragraph_run_carries_text_and_font_kwargs(context):
    paragraph = context.paragraph
    # -- exactly one run was spawned for the `text=` argument --
    assert len(paragraph.runs) == 1, "len(paragraph.runs) == %d" % len(paragraph.runs)
    run = paragraph.runs[0]
    assert run.text == "styled line", "run.text == %r" % run.text
    assert run.font.italic is True, "run.font.italic is %r" % run.font.italic
    assert run.font.name == "Arial", "run.font.name == %r" % run.font.name


# -- _Paragraph.write_rich (#753) -------------------------


@when("I call paragraph.write_rich with a mix of plain, bold, and sized parts")
def when_I_call_paragraph_write_rich_mixed(context):
    # -- start with an empty paragraph so only the runs we author are present --
    context.paragraph.clear()
    context.paragraph.write_rich(
        "plain ",
        ("bold ", {"bold": True}),
        ("big red", {"size": Pt(24), "color": RGBColor(0xFF, 0x00, 0x00)}),
    )


@then("the paragraph contains three runs with the expected text and formatting")
def then_paragraph_has_three_runs(context):
    runs = context.paragraph.runs
    assert len(runs) == 3, "expected 3 runs, got %d" % len(runs)

    assert runs[0].text == "plain ", "runs[0].text == %r" % runs[0].text
    assert runs[0].font.bold is None, "runs[0].font.bold == %r" % runs[0].font.bold

    assert runs[1].text == "bold ", "runs[1].text == %r" % runs[1].text
    assert runs[1].font.bold is True, "runs[1].font.bold == %r" % runs[1].font.bold

    assert runs[2].text == "big red", "runs[2].text == %r" % runs[2].text
    assert runs[2].font.size == Pt(24), "runs[2].font.size == %r" % runs[2].font.size
    assert runs[2].font.color.rgb == RGBColor(0xFF, 0x00, 0x00), (
        "runs[2].font.color.rgb == %r" % runs[2].font.color.rgb
    )


# -- _Paragraph.clone_from / _Run.clone_from (CLO-6) ---------------------


def _build_two_text_frame_fixture(context):
    """Populate `context.src_tf` and `context.dst_tf` in a fresh presentation."""
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    src_tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    dst_tb = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
    context.src_tf = src_tb.text_frame
    context.dst_tf = dst_tb.text_frame


@given("a source paragraph with bullet, level, alignment, and a formatted run")
def given_source_paragraph_with_pPr_and_formatted_run(context):
    _build_two_text_frame_fixture(context)
    src_p = context.src_tf.paragraphs[0]
    src_p.alignment = PP_ALIGN.CENTER
    src_p.level = 2
    src_p.bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD, start_at=3)
    run = src_p.add_run()
    run.text = "Hello"
    run.font.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(0xFF, 0x66, 0x00)
    context.src_paragraph = src_p


@given("a destination paragraph in a separate text frame")
def given_destination_paragraph_in_separate_text_frame(context):
    context.dst_paragraph = context.dst_tf.paragraphs[0]


@when("I call dst_paragraph.clone_from(src_paragraph)")
def when_call_dst_paragraph_clone_from_src(context):
    context.dst_paragraph.clone_from(context.src_paragraph)


@when("I call dst_paragraph.clone_from(src_paragraph) capturing the return")
def when_call_dst_paragraph_clone_from_src_capturing(context):
    context.return_value = context.dst_paragraph.clone_from(context.src_paragraph)


@then("dst_paragraph.level equals src_paragraph.level")
def then_dst_paragraph_level_equals_src(context):
    assert context.dst_paragraph.level == context.src_paragraph.level, (
        "dst level %r != src level %r"
        % (context.dst_paragraph.level, context.src_paragraph.level)
    )


@then("dst_paragraph.alignment equals src_paragraph.alignment")
def then_dst_paragraph_alignment_equals_src(context):
    assert context.dst_paragraph.alignment == context.src_paragraph.alignment


@then("dst_paragraph.bullet.type equals src_paragraph.bullet.type")
def then_dst_paragraph_bullet_type_equals_src(context):
    assert context.dst_paragraph.bullet.type == context.src_paragraph.bullet.type


@then("dst_paragraph.runs carry the same text and formatting as the source")
def then_dst_paragraph_runs_match_src(context):
    src_runs = context.src_paragraph.runs
    dst_runs = context.dst_paragraph.runs
    assert len(dst_runs) == len(src_runs), (
        "len(dst.runs) == %d != %d" % (len(dst_runs), len(src_runs))
    )
    for src_r, dst_r in zip(src_runs, dst_runs):
        assert dst_r.text == src_r.text
        assert dst_r.font.bold == src_r.font.bold
        assert dst_r.font.size == src_r.font.size
        assert dst_r.font.color.rgb == src_r.font.color.rgb


@then("the return value is the destination paragraph")
def then_return_value_is_destination_paragraph(context):
    assert context.return_value is context.dst_paragraph


@given("a source run with bold, italic, size, color, and highlight")
def given_source_run_with_formatting_and_highlight(context):
    _build_two_text_frame_fixture(context)
    src_p = context.src_tf.paragraphs[0]
    run = src_p.add_run()
    run.text = "SRC"
    run.font.bold = True
    run.font.italic = True
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
    run.font.highlight_color.rgb = RGBColor(0xFF, 0xFF, 0x00)
    context.src_run = run


@given("a destination run in a separate text frame")
def given_destination_run_in_separate_text_frame(context):
    dst_p = context.dst_tf.paragraphs[0]
    run = dst_p.add_run()
    run.text = "orig"
    context.dst_run = run


@when("I call dst_run.clone_from(src_run)")
def when_call_dst_run_clone_from_src(context):
    context.dst_run.clone_from(context.src_run)


@when("I call dst_run.clone_from(src_run) capturing the return")
def when_call_dst_run_clone_from_src_capturing(context):
    context.return_value = context.dst_run.clone_from(context.src_run)


@then("dst_run has the same text and formatting as src_run")
def then_dst_run_has_same_text_and_formatting(context):
    src = context.src_run
    dst = context.dst_run
    assert dst.text == src.text, "dst.text %r != src.text %r" % (dst.text, src.text)
    assert dst.font.bold == src.font.bold
    assert dst.font.italic == src.font.italic
    assert dst.font.size == src.font.size
    assert dst.font.color.rgb == src.font.color.rgb
    assert dst.font.highlight_color.rgb == src.font.highlight_color.rgb


@then("the return value is the destination run")
def then_return_value_is_destination_run(context):
    assert context.return_value is context.dst_run
