# pyright: reportPrivateUsage=false

"""Verify-close regression test for issue #116 — ``Picture.replace_image``.

Issue #116 (https://github.com/scanny/python-pptx/issues/116) asked for
a first-class way to swap the pixel bytes of a picture shape without
otherwise disturbing it. The thread accumulated 23 comments over the
course of several years — a reliable signal of heavy demand — with
users posting brittle workarounds that poked at :class:`.ImagePart`
internals, deleted-and-re-added pictures (losing position, name, and
``cNvPr/@descr`` in the process) or attempted to monkey-patch
``Image._blob`` directly.

The supported API :meth:`.Picture.replace_image` was shipped per
``FEATURES.md`` in the CalVer 2026.05.0 release line. Its contract,
per the method docstring, is to preserve every shape-level attribute
while replacing only the pixel bytes: "Position, size, rotation,
cropping, masking shape, outline, and any other shape-level
formatting are preserved — only the pixel bytes of the image are
replaced."

This suite is the *verify-close* pass for #116 — a breadth-first
regression test that pins each preservation guarantee against silent
breakage. Each scenario exercises the real public API from the caller's
perspective, uses no mocks, and where shape-level formatting is under
test, includes a full save/reopen round-trip so the assertion covers
the serialized XML and not just the in-memory proxy.

Sister tests #834 (end-to-end geometry + crop round-trip) and #819
(find-by-alt-text workflow) cover partially overlapping ground; this
file is the catalogue the next regression candidate can be matched
against.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches, Pt


class DescribeIssue116ReplaceImage:
    """Verify-close regression suite for Picture.replace_image (issue #116)."""

    def it_swaps_image_bytes_given_a_file_path(self):
        # -- the canonical #116 flow: pass a filesystem path and expect
        # -- the embedded blob to become the new file's bytes verbatim.
        prs, picture = _prs_with_picture("tests/test_files/python-powered.png")
        original_blob = picture.image.blob

        picture.replace_image("tests/test_files/monty-truth.png")

        with open("tests/test_files/monty-truth.png", "rb") as f:
            expected = f.read()
        assert picture.image.blob == expected
        assert picture.image.blob != original_blob

    def it_swaps_image_bytes_given_a_BytesIO_stream(self):
        # -- #116 commenters repeatedly asked for file-like input so bytes
        # -- produced in-memory (PIL, a download, a thumbnailer) need not
        # -- hit disk first; replace_image must accept any readable stream.
        prs, picture = _prs_with_picture("tests/test_files/python-powered.png")
        with open("tests/test_files/monty-truth.png", "rb") as f:
            new_bytes = f.read()

        picture.replace_image(io.BytesIO(new_bytes))

        assert picture.image.blob == new_bytes

    def it_preserves_position_left_and_top_across_the_replace(self):
        # -- the delete-and-re-add workaround that replace_image supplants
        # -- forces the caller to capture and restore left/top. The whole
        # -- point of the API is that those coordinates are preserved by
        # -- the library — pin that across save/reopen.
        prs, picture = _prs_with_picture(
            "tests/test_files/python-powered.png",
            x=Inches(1.5),
            y=Inches(2.25),
            cx=Inches(2),
            cy=Inches(2),
        )

        picture.replace_image("tests/test_files/monty-truth.png")

        reloaded = _roundtrip_first_picture(prs)
        assert reloaded.left == Inches(1.5)
        assert reloaded.top == Inches(2.25)

    def it_preserves_size_width_and_height_across_the_replace(self):
        # -- same rationale as position: the reporter of #116 specifically
        # -- wanted to avoid recomputing cx/cy when swapping image bytes.
        prs, picture = _prs_with_picture(
            "tests/test_files/python-powered.png",
            x=Emu(0),
            y=Emu(0),
            cx=Inches(4),
            cy=Inches(3),
        )

        picture.replace_image("tests/test_files/monty-truth.png")

        reloaded = _roundtrip_first_picture(prs)
        assert reloaded.width == Inches(4)
        assert reloaded.height == Inches(3)

    def it_preserves_all_four_crop_fractions_across_the_replace(self):
        # -- cropping rectangles live on p:blipFill/a:srcRect, the same
        # -- parent element as the a:blip that replace_image rebinds. The
        # -- rebind must not disturb its sibling; pin all four edges.
        prs, picture = _prs_with_picture("tests/test_files/python-powered.png")
        picture.crop_left = 0.10
        picture.crop_right = 0.20
        picture.crop_top = 0.05
        picture.crop_bottom = 0.15

        picture.replace_image("tests/test_files/monty-truth.png")

        reloaded = _roundtrip_first_picture(prs)
        assert abs(reloaded.crop_left - 0.10) < 1e-6
        assert abs(reloaded.crop_right - 0.20) < 1e-6
        assert abs(reloaded.crop_top - 0.05) < 1e-6
        assert abs(reloaded.crop_bottom - 0.15) < 1e-6

    def it_preserves_alt_text_description_and_title_across_the_replace(self):
        # -- #116's most-upvoted comment chain wanted to locate shapes by
        # -- alt-text across runs and swap them in place. Both cNvPr/@descr
        # -- (alt_text) and cNvPr/@title must survive intact — see #819 for
        # -- the end-to-end find-by-descr workflow this guarantees.
        prs, picture = _prs_with_picture("tests/test_files/python-powered.png")
        picture.alt_text = "Company logo — hero banner"
        picture.title = "Company logo"

        picture.replace_image("tests/test_files/monty-truth.png")

        reloaded = _roundtrip_first_picture(prs)
        assert reloaded.alt_text == "Company logo — hero banner"
        assert reloaded.title == "Company logo"

    def it_preserves_the_shape_name_across_the_replace(self):
        # -- cNvPr/@name is the key used by SlideShapes.get_by_name; a
        # -- replace-image that scrambled it would silently break name-
        # -- based lookups in downstream workflows.
        prs, picture = _prs_with_picture("tests/test_files/python-powered.png")
        picture.name = "LogoPlaceholder"

        picture.replace_image("tests/test_files/monty-truth.png")

        reloaded = _roundtrip_first_picture(prs)
        assert reloaded.name == "LogoPlaceholder"

    def it_preserves_the_picture_outline_across_the_replace(self):
        # -- the method docstring promises outline preservation. Outline
        # -- lives on p:pic/p:spPr/a:ln, a different subtree from the
        # -- a:blipFill that replace_image touches — but this pins the
        # -- guarantee against any future refactor that walks the whole
        # -- spPr looking for formatting fragments.
        from pptx.dml.color import RGBColor

        prs, picture = _prs_with_picture("tests/test_files/python-powered.png")
        picture.line.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        picture.line.width = Pt(3)

        picture.replace_image("tests/test_files/monty-truth.png")

        reloaded = _roundtrip_first_picture(prs)
        assert reloaded.line.color.rgb == RGBColor(0xFF, 0x00, 0x00)
        assert reloaded.line.width == Pt(3)

    def it_preserves_the_masking_auto_shape_across_the_replace(self):
        # -- a picture can be masked by an autoshape (p:pic/p:spPr/a:prstGeom).
        # -- #116 explicitly names "masking shape" as a preserved attribute.
        # -- Use OVAL so the assertion catches any drop-back to RECTANGLE.
        prs, picture = _prs_with_picture("tests/test_files/python-powered.png")
        picture.auto_shape_type = MSO_SHAPE.OVAL
        assert picture.auto_shape_type == MSO_SHAPE.OVAL

        picture.replace_image("tests/test_files/monty-truth.png")

        reloaded = _roundtrip_first_picture(prs)
        assert reloaded.auto_shape_type == MSO_SHAPE.OVAL

    def it_preserves_rotation_across_the_replace(self):
        # -- rotation is an xfrm/@rot attribute on the p:pic spPr; not
        # -- directly listed in the docstring's preservation list but
        # -- implied by "any other shape-level formatting". Round-trip
        # -- to confirm the in-memory and serialized values agree.
        prs, picture = _prs_with_picture("tests/test_files/python-powered.png")
        picture.rotation = 30.0

        picture.replace_image("tests/test_files/monty-truth.png")

        reloaded = _roundtrip_first_picture(prs)
        assert reloaded.rotation == 30.0

    def it_survives_a_full_save_and_reopen_round_trip(self):
        # -- the end-to-end confidence test: stamp every preservation
        # -- attribute the docstring promises, save, reopen, and re-
        # -- verify each one. This is the single scenario the next
        # -- maintainer is likely to sanity-check when triaging a new
        # -- issue that smells like #116.
        from pptx.dml.color import RGBColor

        prs, picture = _prs_with_picture(
            "tests/test_files/python-powered.png",
            x=Inches(1),
            y=Inches(2),
            cx=Inches(3),
            cy=Inches(4),
        )
        picture.name = "RoundTripPicture"
        picture.alt_text = "round-trip alt text"
        picture.title = "round-trip title"
        picture.rotation = 15.0
        picture.crop_left = 0.05
        picture.crop_right = 0.07
        picture.crop_top = 0.09
        picture.crop_bottom = 0.11
        picture.auto_shape_type = MSO_SHAPE.ROUNDED_RECTANGLE
        picture.line.color.rgb = RGBColor(0x12, 0x34, 0x56)
        picture.line.width = Pt(2)

        picture.replace_image("tests/test_files/monty-truth.png")

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        reloaded = prs2.slides[0].shapes[0]

        with open("tests/test_files/monty-truth.png", "rb") as f:
            assert reloaded.image.blob == f.read()
        assert reloaded.name == "RoundTripPicture"
        assert reloaded.alt_text == "round-trip alt text"
        assert reloaded.title == "round-trip title"
        assert reloaded.left == Inches(1)
        assert reloaded.top == Inches(2)
        assert reloaded.width == Inches(3)
        assert reloaded.height == Inches(4)
        assert reloaded.rotation == 15.0
        assert abs(reloaded.crop_left - 0.05) < 1e-6
        assert abs(reloaded.crop_right - 0.07) < 1e-6
        assert abs(reloaded.crop_top - 0.09) < 1e-6
        assert abs(reloaded.crop_bottom - 0.11) < 1e-6
        assert reloaded.auto_shape_type == MSO_SHAPE.ROUNDED_RECTANGLE
        assert reloaded.line.color.rgb == RGBColor(0x12, 0x34, 0x56)
        assert reloaded.line.width == Pt(2)

    def it_rejects_a_picture_with_no_embedded_image(self):
        # -- the one documented ValueError branch of the API contract.
        # -- Constructing a p:pic with a:blip but no r:embed exercises
        # -- the guard without needing a purpose-built fixture file.
        from pptx.shapes.picture import Picture

        from .unitutil.cxml import element

        pic = element("p:pic/p:blipFill/a:blip")
        picture = Picture(pic, None)

        with pytest.raises(ValueError, match="no embedded image to replace"):
            picture.replace_image("tests/test_files/monty-truth.png")


# -- helpers -----------------------------------------------------------


def _prs_with_picture(
    image_path,
    x=Inches(1),
    y=Inches(1),
    cx=Inches(2),
    cy=Inches(2),
):
    """Return ``(prs, picture)`` for a fresh single-slide presentation."""
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    picture = slide.shapes.add_picture(image_path, x, y, cx, cy)
    return prs, picture


def _roundtrip_first_picture(prs):
    """Save `prs` to an in-memory buffer, reopen, and return its first picture.

    Using this for every preservation-oriented scenario means the assertion
    covers the serialized XML rather than just the in-memory proxy — which
    is where `replace_image` bugs would most plausibly hide.
    """
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf).slides[0].shapes[0]
