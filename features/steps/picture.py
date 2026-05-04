"""Gherkin step implementations for picture-related features."""

from __future__ import annotations

import io

from behave import given, then, when
from helpers import saved_pptx_path, test_image, test_pptx

from pptx import Presentation
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.package import Package
from pptx.parts.image import Image
from pptx.util import Inches

# given ===================================================


@given("a Picture object masked by a {shape} as picture")
def given_a_picture_object_masked_by_shape_as_picture(context, shape):
    shape_idx = {"rectangle": 0, "circle": 1}[shape]
    prs = Presentation(test_pptx("shp-picture"))
    context.picture = prs.slides[1].shapes[shape_idx]


@given("a picture of known position and size")
def given_a_picture_of_known_position_and_size(context):
    prs = Presentation(test_pptx("shp-pos-and-size"))
    context.picture = prs.slides[1].shapes[0]


@given("an Image object loaded from {filename}")
def given_an_Image_object_loaded_from_filename(context, filename):
    context.image = Image.from_file(test_image(filename))


# when ====================================================


@when("I add the image {filename} using shapes.add_picture()")
def when_I_add_the_image_filename_using_shapes_add_picture(context, filename):
    shapes = context.slide.shapes
    shapes.add_picture(test_image(filename), Inches(1.25), Inches(1.25))


@when("I add the stream image {filename} using shapes.add_picture()")
def when_I_add_the_stream_image_filename_using_add_picture(context, filename):
    shapes = context.slide.shapes
    with open(test_image(filename), "rb") as f:
        stream = io.BytesIO(f.read())
    shapes.add_picture(stream, Inches(1.25), Inches(1.25))


@when("I assign MSO_AUTO_SHAPE_TYPE.{member} to picture.auto_shape_type")
def when_I_assign_member_to_picture_auto_shape_type(context, member):
    context.picture.auto_shape_type = getattr(MSO_AUTO_SHAPE_TYPE, member)


@when("I assign {value} to picture.transparency")
def when_I_assign_value_to_picture_transparency(context, value):
    context.picture.transparency = float(value)


@when("I add an SVG picture via shapes.add_picture_svg()")
def when_I_add_an_svg_picture_via_add_picture_svg(context):
    svg_bytes = (
        b'<?xml version="1.0" encoding="UTF-8"?>\n'
        b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">\n'
        b'  <circle cx="50" cy="50" r="40" fill="#4B8BBE"/>\n'
        b"</svg>\n"
    )
    shapes = context.slide.shapes
    context.shape = shapes.add_picture_svg(
        io.BytesIO(svg_bytes), Inches(1), Inches(1), Inches(2), Inches(2)
    )


# then ====================================================


@then("a {ext} image part appears in the pptx file")
def step_then_a_ext_image_part_appears_in_the_pptx_file(context, ext):
    pkg = Package.open(saved_pptx_path)
    partnames = frozenset(p.partname for p in pkg.iter_parts())
    image_partname = "/ppt/media/image1.%s" % ext
    assert image_partname in partnames, "got %s" % [p for p in partnames if "image" in p]


@then("picture.auto_shape_type == MSO_AUTO_SHAPE_TYPE.{member}")
def then_picture_auto_shape_type_eq_shape_type_member(context, member):
    expected = getattr(MSO_AUTO_SHAPE_TYPE, member)
    actual = context.picture.auto_shape_type
    assert actual == expected, "shape.auto_shape_type == %s" % actual


@then("the picture appears in the slide")
def then_the_picture_appears_in_the_slide(context):
    prs = Presentation(saved_pptx_path)
    slide = prs.slides[0]
    shapes = slide.shapes
    cls_names = [sp.__class__.__name__ for sp in shapes]
    assert "Picture" in cls_names


@then('image.ext == "{expected}"')
def then_image_ext_eq(context, expected):
    actual = context.image.ext
    assert actual == expected, "image.ext == %r" % actual


@then('image.content_type == "{expected}"')
def then_image_content_type_eq(context, expected):
    actual = context.image.content_type
    assert actual == expected, "image.content_type == %r" % actual


@then("picture.transparency == {expected}")
def then_picture_transparency_eq(context, expected):
    expected_value = float(expected)
    actual = context.picture.transparency
    assert abs(actual - expected_value) < 1e-6, "picture.transparency == %r" % actual


@then("an SVG media part appears in the pptx file")
def then_an_svg_media_part_appears(context):
    pkg = Package.open(saved_pptx_path)
    partnames = [str(p.partname) for p in pkg.iter_parts()]
    svg_parts = [p for p in partnames if p.endswith(".svg")]
    assert svg_parts, "no SVG part found; got %r" % partnames


@then("the blipFill a:blip carries an asvg:svgBlip extension")
def then_the_blipfill_blip_carries_svgBlip_extension(context):
    prs = Presentation(saved_pptx_path)
    slide = prs.slides[0]
    pic = next(sp for sp in slide.shapes if sp.__class__.__name__ == "Picture")
    blips = pic._element.xpath(  # noqa: SLF001
        "./p:blipFill/a:blip"
    )
    assert blips, "no a:blip element found"
    svgBlips = blips[0].xpath(".//asvg:svgBlip")
    assert svgBlips, "asvg:svgBlip extension missing from a:blip"
    # -- the svgBlip carries r:embed pointing at the SVG part --
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rId = svgBlips[0].get("{%s}embed" % r_ns)
    assert rId, "asvg:svgBlip missing r:embed"
