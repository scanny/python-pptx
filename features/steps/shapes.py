"""Gherkin step implementations for shape collections."""

from __future__ import annotations

import io
import zipfile

from behave import given, then, when
from helpers import saved_pptx_path, test_file, test_image, test_pptx

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER, PROG_ID
from pptx.shapes.base import BaseShape
from pptx.util import Emu, Inches

# given ===================================================


@given("a _BaseShapes object as shapes")
def given_a_BaseShapes_object_as_shapes(context):
    prs = Presentation()
    context.shapes = prs.slides.add_slide(prs.slide_layouts[6]).shapes


@given("a GroupShapes object as shapes")
@given("a GroupShapes object of length 3 as shapes")
def given_a_GroupShapes_object_of_length_3_as_shapes(context):
    prs = Presentation(test_pptx("shp-groupshape"))
    group_shape = prs.slides[0].shapes[0]
    context.shapes = group_shape.shapes


@given("a LayoutPlaceholders object of length 2 as shapes")
def given_a_LayoutPlaceholders_object_of_length_2_as_shapes(context):
    prs = Presentation(test_pptx("lyt-shapes"))
    context.shapes = prs.slide_layouts[0].placeholders


@given("a LayoutShapes object of length 3 as shapes")
def given_a_LayoutShapes_object_of_length_3_as_shapes(context):
    prs = Presentation(test_pptx("lyt-shapes"))
    context.shapes = prs.slide_layouts[0].shapes


@given("a MasterPlaceholders object of length 2 as shapes")
def given_a_MasterPlaceholders_object_of_length_2_as_shapes(context):
    prs = Presentation(test_pptx("mst-placeholders"))
    context.shapes = prs.slide_masters[0].placeholders


@given("a MasterShapes object of length 2 as shapes")
def given_a_MasterShapes_object_of_length_2_as_shapes(context):
    prs = Presentation(test_pptx("mst-shapes"))
    context.shapes = prs.slide_masters[0].shapes


@given("a {PROG_ID_member} file as ole_object_file")
def given_a_PROG_ID_member_file_as_ole_object_file(context, PROG_ID_member):
    filename = {
        "DOCX": "shp-embedded-docx.docx",
        "PPTX": "shp-embedded-pptx.pptx",
        "XLSX": "shp-embedded-xlsx.xlsx",
    }[PROG_ID_member]
    with open(test_file(filename), "rb") as f:
        context.ole_object_file = io.BytesIO(f.read())
    context.PROG_ID_member = PROG_ID_member


@given("an in-memory zip archive as ole_object_file")
def given_an_in_memory_zip_archive_as_ole_object_file(context):
    """Build a tiny zip archive in memory to exercise the generic-embed path (#752)."""
    import zipfile

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("hello.txt", b"hello world")
    context.ole_object_file = io.BytesIO(buf.getvalue())


@given("a SlidePlaceholders object of length 2 as shapes")
def given_a_SlidePlaceholders_object_of_length_2_as_shapes(context):
    prs = Presentation(test_pptx("shp-shapes"))
    context.shapes = prs.slides[0].placeholders


@given("a SlideShapes object as shapes")
def given_a_SlideShapes_object_as_shapes(context):
    prs = Presentation(test_pptx("shp-shapes"))
    context.shapes = prs.slides[0].shapes


@given("a SlideShapes object containing {a_or_no} movies")
def given_a_SlideShapes_object_containing_a_or_no_movies(context, a_or_no):
    pptx = {"one or more": "shp-movie-props", "no": "shp-shapes"}[a_or_no]
    prs = Presentation(test_pptx(pptx))
    context.prs = prs
    context.shapes = prs.slides[0].shapes


@given(
    "a SlideShapes object whose slide has an "
    "mc:AlternateContent-wrapped p:timing"
)
def given_a_SlideShapes_obj_whose_slide_has_mc_wrapped_timing(context):
    """Regression fixture for issue #954.

    Builds an in-memory deck whose first slide carries a pre-existing
    ``p:timing`` element wrapped inside an
    ``mc:AlternateContent``/``mc:Choice`` block (the form PowerPoint
    emits when the timing references 2010+ extensions such as a morph
    trigger). Before the fix, a subsequent ``add_movie()`` call
    produced a second, orphan ``p:timing`` sibling of the wrapper.
    """
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # --- inject the wrapped timing onto the p:sld element ---
    sld = slide._element
    wrapper_xml = (
        "<mc:AlternateContent %s>\n"
        '  <mc:Choice xmlns:p14="http://schemas.microsoft.com/office/po'
        'werpoint/2010/main" Requires="p14">\n'
        "    <p:timing>\n"
        "      <p:tnLst>\n"
        "        <p:par>\n"
        '          <p:cTn id="1" dur="indefinite" restart="never" '
        'nodeType="tmRoot">\n'
        "            <p:childTnLst/>\n"
        "          </p:cTn>\n"
        "        </p:par>\n"
        "      </p:tnLst>\n"
        "    </p:timing>\n"
        "  </mc:Choice>\n"
        "  <mc:Fallback>\n"
        "    <p:timing>\n"
        "      <p:tnLst>\n"
        "        <p:par>\n"
        '          <p:cTn id="1" dur="indefinite" restart="never" '
        'nodeType="tmRoot"/>\n'
        "        </p:par>\n"
        "      </p:tnLst>\n"
        "    </p:timing>\n"
        "  </mc:Fallback>\n"
        "</mc:AlternateContent>" % nsdecls("p", "mc")
    )
    sld.append(parse_xml(wrapper_xml))
    context.prs = prs
    context.slide = slide
    context.shapes = slide.shapes


@given("a SlideShapes object from a saved deck already containing a wav audio")
def given_a_SlideShapes_object_from_a_saved_deck_with_wav_audio(context):
    # -- author a fresh deck that contains a wav audio movie shape, save it to
    # -- an in-memory buffer, then reopen it so existing audio parts are loaded
    # -- from-file (this matches the issue #926 reproduction) and expose the
    # -- first slide's shapes for the subsequent add_movie(...) call --
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    x, y, cx, cy = Emu(914400), Emu(914400), Emu(914400), Emu(914400)
    slide.shapes.add_movie(
        test_file("silence.wav"), x, y, cx, cy, mime_type="audio/x-wav"
    )
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    reopened = Presentation(buf)
    context.prs = reopened
    context.shapes = reopened.slides[0].shapes


@when("I call shapes.add_movie(audio_file, mime_type='audio/x-wav')")
def when_I_call_shapes_add_movie_x_wav(context):
    shapes = context.shapes
    x, y, cx, cy = Emu(2590800), Emu(571500), Emu(914400), Emu(914400)
    context.movie = shapes.add_movie(
        test_file("silence.wav"), x, y, cx, cy, mime_type="audio/x-wav"
    )


@given("a SlideShapes object of length 6 shapes as shapes")
def given_a_SlideShapes_object_of_length_6_as_shapes(context):
    prs = Presentation(test_pptx("shp-shapes"))
    context.shapes = prs.slides[0].shapes


@given("a SlideShapes object having a {type} shape at offset {idx}")
def given_a_SlideShapes_obj_having_type_shape_at_off_idx(context, type, idx):
    prs = Presentation(test_pptx("shp-shapes"))
    context.shapes = prs.slides[1].shapes


@given("a slide with an mc:AlternateContent-wrapped shape")
def given_a_slide_with_an_mc_AlternateContent_wrapped_shape(context):
    context.prs = Presentation(test_pptx("shp-mc-alternate-content"))
    context.slide = context.prs.slides[0]


@given("a slide with a math-equation shape")
def given_a_slide_with_a_math_equation_shape(context):
    context.prs = Presentation(test_pptx("shp-math-equation"))
    context.slide = context.prs.slides[0]


@given("a Presentation whose slide-layout owns an audio/mpeg part")
def given_a_Presentation_whose_layout_owns_an_mp3(context):
    """Build a pptx package where ``slideLayout1`` carries an ``audio/mpeg`` rel.

    Regression fixture for issue #323 — before #502/#734 a subsequent
    ``add_movie()`` call on the loaded deck raised ``AttributeError`` inside
    ``_MediaParts._find_by_sha1``.
    """
    src = test_pptx("shp-shapes")
    with open(src, "rb") as f:
        src_bytes = f.read()

    with zipfile.ZipFile(io.BytesIO(src_bytes), "r") as zin:
        items = {name: zin.read(name) for name in zin.namelist()}

    items["ppt/media/media1.mp3"] = b"ID3\x03\x00\x00\x00\x00\x00\x00mp3-bytes"

    ct_xml = items["[Content_Types].xml"].decode("utf-8")
    ct_xml = ct_xml.replace(
        "</Types>",
        '<Override PartName="/ppt/media/media1.mp3"'
        ' ContentType="audio/mpeg"/></Types>',
    )
    items["[Content_Types].xml"] = ct_xml.encode("utf-8")

    rels_key = "ppt/slideLayouts/_rels/slideLayout1.xml.rels"
    rels_xml = items.get(rels_key, b"").decode("utf-8") or (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/'
        'package/2006/relationships"/>'
    )
    rels_xml = rels_xml.replace(
        "</Relationships>",
        '<Relationship Id="rIdAudio1"'
        ' Type="http://schemas.microsoft.com/office/2007/relationships/media"'
        ' Target="../media/media1.mp3"/></Relationships>',
    )
    items[rels_key] = rels_xml.encode("utf-8")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)

    context.prs = Presentation(io.BytesIO(buf.getvalue()))


# when ====================================================


@when("I add 100 shapes")
def when_I_add_100_shapes(context):
    X_ORIG = Y_ORIG = Inches(0.0625)
    X_INCR = Y_INCR = Inches(0.5)
    CX = CY = Inches(0.375)

    def iter_corner():
        y = Y_ORIG
        while True:
            for i in range(20):
                x = X_ORIG + (X_INCR * i)
                yield x, y
            y += Y_INCR

    shapes = context.shapes
    corners = iter_corner()
    for i in range(100):
        x, y = next(corners)
        shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, CX, CY)


@when("I add a table to the slide's shape collection")
def when_I_call_shapes_add_table(context):
    shapes = context.slide.shapes
    x, y = (Inches(1.00), Inches(2.00))
    cx, cy = (Inches(3.00), Inches(1.00))
    shapes.add_table(2, 2, x, y, cx, cy)


@when("I add a table to the slide's shape collection using float dimensions")
def when_I_call_shapes_add_table_with_floats(context):
    # -- regression for #288: `add_table()` must accept float position/size arguments produced
    # -- by arithmetic like `0.1 * slide_width`.
    shapes = context.slide.shapes
    x, y = (Inches(1.00) * 0.5, Inches(2.00) * 0.5)
    cx, cy = (Inches(3.00) * 0.5, Inches(1.00) * 0.5)
    assert isinstance(x, float)
    assert isinstance(cx, float)
    shapes.add_table(2, 3, x, y, cx, cy)


@when("I assign shape.ole_format to ole_format")
def when_I_assign_shape_ole_format_to_ole_format(context):
    context.ole_format = context.shape.ole_format


@when("I assign shapes.add_chart() to shape")
def when_I_assign_shapes_add_chart_to_shape(context):
    chart_data = CategoryChartData()
    chart_data.categories = ("Foo", "Bar")
    chart_data.add_series("East", (1.0, 2.0))
    chart_data.add_series("West", (3.0, 4.0))

    context.shape = context.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(8),
        Inches(5),
        chart_data,
    )


@when("I assign shapes.add_connector() to shape")
def when_I_assign_shapes_add_connector_to_shape(context):
    context.shape = context.shapes.add_connector(MSO_CONNECTOR.CURVE, 4, 3, 2, 1)


@when("I assign shapes.add_group_shape() to shape")
def when_I_assign_shapes_add_group_shape_to_shape(context):
    context.shape = context.shapes.add_group_shape()


@when("I assign shapes.add_ole_object(ole_object_file) to shape")
def when_I_assign_shapes_add_ole_object_to_shape(context):
    context.shape = context.shapes.add_ole_object(
        context.ole_object_file, getattr(PROG_ID, context.PROG_ID_member), 4, 3, 2, 1
    )


@when(
    'I assign shapes.add_ole_object(ole_object_file, "{prog_id}", "{extension}") to shape'
)
def when_I_assign_shapes_add_ole_object_generic_to_shape(context, prog_id, extension):
    """Exercise the arbitrary-progId path added in #752."""
    context.shape = context.shapes.add_ole_object(
        context.ole_object_file,
        prog_id,
        4,
        3,
        2,
        1,
        extension=extension,
    )


@when("I assign shapes.add_picture() to shape")
def when_I_assign_shapes_add_picture_to_shape(context):
    context.shape = context.shapes.add_picture(test_image("sonic.gif"), Inches(1), Inches(2))


@when("I assign shapes.add_shape() to shape")
def when_I_assign_shapes_add_shape_to_shape(context):
    context.shape = context.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2), Inches(3), Inches(1), Inches(0.5)
    )


@when("I assign shapes.add_textbox() to shape")
def when_I_assign_shapes_add_textbox_to_shape(context):
    context.shape = context.shapes.add_textbox(Inches(1), Inches(2), Inches(3), Inches(0.5))


@when("I assign shapes.build_freeform() to builder")
def when_I_assign_shapes_build_freeform_to_builder(context):
    shapes = context.shapes
    builder = shapes.build_freeform()
    context.builder = builder


@when("I assign shapes.build_freeform(scale=100.0) to builder")
def when_I_assign_shapes_build_freeform_scale_to_builder(context):
    shapes = context.shapes
    builder = shapes.build_freeform(scale=100.0)
    context.builder = builder


@when("I assign shapes.build_freeform(scale=(200.0, 100.0)) to builder")
def when_I_assign_shapes_build_freeform_scale_rectnglr_to_builder(context):
    shapes = context.shapes
    builder = shapes.build_freeform(scale=(200.0, 100.0))
    context.builder = builder


@when("I assign shapes.build_freeform(start_x=25, start_y=125) to builder")
def when_I_assign_shapes_build_freeform_start_x_start_y_to_builder(context):
    shapes = context.shapes
    builder = shapes.build_freeform(25, 125)
    context.builder = builder


@when("I assign True to shapes.turbo_add_enabled")
def when_I_assign_True_to_shapes_turbo_add_enabled(context):
    context.shapes.turbo_add_enabled = True


@when("I call shapes.add_chart({type_}, chart_data)")
def when_I_call_shapes_add_chart(context, type_):
    chart_type = getattr(XL_CHART_TYPE, type_)
    context.chart = context.shapes.add_chart(chart_type, 0, 0, 0, 0, context.chart_data).chart


@when("I call shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 1, 2, 3, 4)")
def when_I_call_shapes_add_connector(context):
    context.connector = context.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 1, 2, 3, 4)


@when("I call shapes.add_movie(file, x, y, cx, cy, poster_frame)")
def when_I_call_shapes_add_movie(context):
    shapes = context.shapes
    x, y, cx, cy = Emu(2590800), Emu(571500), Emu(3962400), Emu(5715000)
    context.movie = shapes.add_movie(
        test_file("just-two-mice.mp4"), x, y, cx, cy, test_file("just-two-mice.png")
    )


@when("I call shapes.add_movie(audio_file, mime_type='audio/wav')")
def when_I_call_shapes_add_movie_audio(context):
    shapes = context.shapes
    x, y, cx, cy = Emu(2590800), Emu(571500), Emu(914400), Emu(914400)
    context.movie = shapes.add_movie(
        test_file("silence.wav"), x, y, cx, cy, mime_type="audio/wav"
    )


@when("I call shapes.add_movie() on a newly added slide")
def when_I_call_shapes_add_movie_on_a_new_slide(context):
    slide = context.prs.slides.add_slide(context.prs.slide_layouts[0])
    context.shapes = slide.shapes
    x, y, cx, cy = Emu(2590800), Emu(571500), Emu(3962400), Emu(5715000)
    context.movie = context.shapes.add_movie(
        test_file("just-two-mice.mp4"),
        x,
        y,
        cx,
        cy,
        test_file("just-two-mice.png"),
    )


# -- issue #427: user-facing autoplay convenience kwarg on add_movie --
@when("I call shapes.add_movie(file, x, y, cx, cy, poster_frame, autoplay=True)")
def when_I_call_shapes_add_movie_autoplay_true(context):
    shapes = context.shapes
    x, y, cx, cy = Emu(2590800), Emu(571500), Emu(3962400), Emu(5715000)
    context.movie = shapes.add_movie(
        test_file("just-two-mice.mp4"),
        x,
        y,
        cx,
        cy,
        test_file("just-two-mice.png"),
        autoplay=True,
    )


# then ====================================================


@then("iterating shapes produces {count} objects of type {class_name}")
def then_iterating_shapes_produces_count_objects_of_type_class_name(context, count, class_name):
    shapes = context.shapes
    expected_count, expected_class_name = int(count), class_name
    idx = -1
    for idx, shape in enumerate(shapes):
        actual_class_name = shape.__class__.__name__
        assert actual_class_name == expected_class_name, (
            "shape.__class__.__name__ == %s" % actual_class_name
        )
    actual_count = idx + 1
    assert actual_count == expected_count, "got %d items" % actual_count


@then("iterating shapes produces {count} objects that subclass BaseShape")
def then_iterating_shapes_produces_count_objects_that_subclass_BaseShape(context, count):
    shapes = context.shapes
    expected_count = int(count)
    idx = -1
    for idx, shape in enumerate(shapes):
        class_name = shape.__class__.__name__
        assert isinstance(shape, BaseShape), "%s does not subclass BaseShape" % class_name
    actual_count = idx + 1
    assert actual_count == expected_count, "got %d items" % actual_count


@then("len(shapes) == {value}")
def then_len_shapes_eq_value(context, value):
    expected_len = int(value)
    actual_len = len(context.shapes)
    assert actual_len == expected_len, "len(shapes) == %s" % actual_len


@then("shape is a {clsname} object")
def then_shape_is_a_type_object(context, clsname):
    actual_class_name = context.shape.__class__.__name__
    expected_class_name = clsname
    assert actual_class_name == expected_class_name, "shape is a %s object" % actual_class_name


@then("shapes[-1] == shape")
def then_shapes_minus_1_eq_shape(context):
    shapes, shape = context.shapes, context.shape
    assert shapes[-1] == shape


@then("shapes[{idx}] is a {type_} object")
def then_shapes_idx_is_a_type_object(context, idx, type_):
    shapes = context.shapes
    type_name = type(shapes[int(idx)]).__name__
    assert type_name == type_, "got %s" % type_name


@then("shapes.get(idx=10) is the body placeholder")
def then_shapes_get_10_is_the_body_placeholder(context):
    shapes = context.shapes
    title_placeholder = shapes.get(idx=0)
    body_placeholder = shapes.get(idx=10)
    assert title_placeholder._element is shapes[0]._element
    assert body_placeholder._element is shapes[1]._element


@then("shapes.get(PP_PLACEHOLDER.BODY) is the body placeholder")
def then_shapes_get_PP_PLACEHOLDER_BODY_is_the_body_ph(context):
    shapes = context.shapes
    title_placeholder = shapes.get(PP_PLACEHOLDER.TITLE)
    body_placeholder = shapes.get(PP_PLACEHOLDER.BODY)
    assert title_placeholder._element is shapes[0]._element
    assert body_placeholder._element is shapes[1]._element


@then("shapes.index(shape) for each shape matches its sequence position")
def then_shapes_index_for_each_shape_matches_sequence_position(context):
    shapes = context.shapes
    for idx, shape in enumerate(shapes):
        assert idx == shapes.index(shape), "index doesn't match for idx == %s" % idx


@then("shapes.title is the title placeholder")
def then_shapes_title_is_the_title_placeholder(context):
    shapes = context.shapes
    title_placeholder = shapes.title
    assert title_placeholder.element is shapes[0].element
    assert title_placeholder.shape_id == 4


@then("shapes.turbo_add_enabled is False")
def then_shapes_turbo_add_enabled_is_False(context):
    shapes = context.shapes
    assert shapes.turbo_add_enabled is False


@then("the table appears in the slide")
def then_the_table_appears_in_the_slide(context):
    prs = Presentation(saved_pptx_path)
    expected_table_graphic_frame = prs.slides[0].shapes[0]
    assert expected_table_graphic_frame.has_table


@then("len(slide.shapes) counts the choice-wrapped shape")
def then_len_slide_shapes_counts_the_choice_wrapped_shape(context):
    assert len(context.slide.shapes) == 3, "expected 3 shapes, got %d" % len(
        context.slide.shapes
    )


@then("slide.shapes surfaces the choice-wrapped shape by name")
def then_slide_shapes_surfaces_the_choice_wrapped_shape_by_name(context):
    names = [shape.name for shape in context.slide.shapes]
    assert "Modern-Choice" in names, "expected 'Modern-Choice' in %r" % names


@then("saving the presentation preserves the mc:Fallback subtree")
def then_saving_the_presentation_preserves_the_mc_Fallback_subtree(context):
    import zipfile

    context.prs.save(saved_pptx_path)
    with zipfile.ZipFile(saved_pptx_path) as z:
        slide_xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
    assert "AlternateContent" in slide_xml, "mc:AlternateContent wrapper not preserved"
    assert "Fallback" in slide_xml, "mc:Fallback subtree not preserved"
    assert "Fallback-Shape" in slide_xml, "mc:Fallback content was stripped on save"
@then("the saved table has integer-valued position and size")
def then_the_saved_table_has_integer_position_and_size(context):
    prs = Presentation(saved_pptx_path)
    graphic_frame = prs.slides[0].shapes[0]
    # -- each position/size value is a Length (int subclass) with no fractional component
    for value in (graphic_frame.left, graphic_frame.top, graphic_frame.width, graphic_frame.height):
        assert isinstance(value, int), value
        assert not isinstance(value, float), value
    # -- every grid-col width and row height in the stored a:tbl XML is also a pure integer str
    tbl = graphic_frame.table._tbl
    for gc in tbl.tblGrid.iterchildren():
        assert gc.get("w").lstrip("-").isdigit(), gc.get("w")
    for tr in tbl.tr_lst:
        assert tr.get("h").lstrip("-").isdigit(), tr.get("h")


def _math_equation_shape(context):
    for shape in context.slide.shapes:
        if shape.name == "Math-Equation":
            return shape
    raise AssertionError("'Math-Equation' shape not found in slide")


@then("shape.has_math_equation is True for the equation shape")
def then_shape_has_math_equation_is_True_for_equation_shape(context):
    shape = _math_equation_shape(context)
    assert shape.has_math_equation is True, "expected has_math_equation to be True"


@then("shape.has_math_equation is False for non-equation shapes")
def then_shape_has_math_equation_is_False_for_non_equation_shapes(context):
    for shape in context.slide.shapes:
        if shape.name == "Math-Equation":
            continue
        assert shape.has_math_equation is False, (
            "expected has_math_equation False for %r" % shape.name
        )


@then("shape.math_equation_xml returns the raw OMML XML")
def then_shape_math_equation_xml_returns_the_raw_OMML_XML(context):
    shape = _math_equation_shape(context)
    oMath_xml = shape.math_equation_xml
    assert oMath_xml is not None, "expected math_equation_xml to be a string"
    assert oMath_xml.startswith("<m:oMath"), "expected OMML to begin with <m:oMath, got %r" % (
        oMath_xml[:40],
    )
    assert (
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"' in oMath_xml
    ), "expected m: namespace declaration in returned OMML"


@then("shape.math_equation_xml is None for non-equation shapes")
def then_shape_math_equation_xml_is_None_for_non_equation_shapes(context):
    for shape in context.slide.shapes:
        if shape.name == "Math-Equation":
            continue
        assert shape.math_equation_xml is None, (
            "expected math_equation_xml None for %r, got %r"
            % (shape.name, shape.math_equation_xml)
        )


@then("the saved presentation reloads without error")
def then_the_saved_presentation_reloads_without_error(context):
    """Regression guard for #323 — confirm the layout-audio pptx round-trips."""
    Presentation(saved_pptx_path)


@then("the slide has exactly one p:timing element")
def then_the_slide_has_exactly_one_p_timing_element(context):
    """Regression guard for issue #954.

    After a ``add_movie()`` call on a slide whose existing ``p:timing``
    is wrapped inside ``mc:AlternateContent``, the slide must still
    contain exactly one ``p:timing`` within the preferred-rendering
    ``mc:Choice`` (the ``mc:Fallback`` copy is ignored) and no second
    ``p:timing`` sibling of the wrapper.
    """
    sld = context.slide._element
    direct_timings = sld.findall(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}timing"
    )
    assert len(direct_timings) == 0, (
        "expected no direct p:timing children of p:sld (timing stays wrapped); "
        "found %d" % len(direct_timings)
    )
    choice_timings = sld.xpath("./mc:AlternateContent/mc:Choice/p:timing")
    assert len(choice_timings) == 1, (
        "expected exactly one p:timing inside mc:Choice; found %d"
        % len(choice_timings)
    )


@then("the movie shape's p:video entry sits inside that p:timing")
def then_the_video_entry_sits_inside_that_timing(context):
    """Regression guard for issue #954 — ensure the new video was merged
    into the pre-existing wrapped `p:timing`, not appended elsewhere."""
    sld = context.slide._element
    videos = sld.xpath("./mc:AlternateContent/mc:Choice/p:timing//p:video")
    assert len(videos) >= 1, "expected a p:video inside the wrapped p:timing"


# --- issue #532: Selection-Pane-equivalent flat traversal -----------------


@given("a slide with 1 top-level shape plus a group of 3 shapes")
def given_a_slide_with_1_top_shape_plus_group_of_3(context):
    # -- `shp-groupshape.pptx` ships with a slide whose shape tree is
    # -- [Group 4 (children: Rounded Rectangle 1, Oval 2, Isosceles
    # -- Triangle 3), Rectangle 5].
    prs = Presentation(test_pptx("shp-groupshape"))
    context.slide = prs.slides[0]
    context.shapes = context.slide.shapes


@then("len(list(slide.shape_tree_flat)) == 5")
def then_len_shape_tree_flat_equals_5(context):
    flat = list(context.slide.shape_tree_flat)
    assert len(flat) == 5, "expected 5 shapes (1 top + 1 group + 3 nested), got %d" % len(
        flat
    )


@then("slide.shape_tree_flat yields the group before its children")
def then_flat_yields_group_before_children(context):
    names = [s.name for s in context.slide.shape_tree_flat]
    # -- group name precedes each nested child in the flat sequence --
    group_idx = names.index("Group 4")
    for child in ("Rounded Rectangle 1", "Oval 2", "Isosceles Triangle 3"):
        assert names.index(child) > group_idx, (
            "expected %r to appear after 'Group 4' in %r" % (child, names)
        )


@then("slide.shapes.descendants() matches slide.shape_tree_flat")
def then_descendants_matches_shape_tree_flat(context):
    a = [s.name for s in context.slide.shape_tree_flat]
    b = [s.name for s in context.slide.shapes.descendants()]
    assert a == b, "shape_tree_flat=%r != descendants=%r" % (a, b)


@then("shapes.get_by_name of an inner-group shape is None by default")
def then_get_by_name_default_miss(context):
    assert context.shapes.get_by_name("Oval 2") is None


@then(
    "shapes.get_by_name of an inner-group shape with include_descendants finds it"
)
def then_get_by_name_with_include_descendants(context):
    hit = context.shapes.get_by_name("Oval 2", include_descendants=True)
    assert hit is not None
    assert hit.name == "Oval 2"
