# pyright: reportPrivateUsage=false

"""Verify-close regression test for issue #419 — picture fill on a shape.

Issue #419 (https://github.com/scanny/python-pptx/issues/419) asked for
a supported way to apply the "Picture or texture fill" option from
PowerPoint's Format Shape pane directly to an auto-shape — the most
common request being to drop an image into a rectangle or oval without
stacking a separate ``Picture`` shape on top. Users on the thread posted
brittle workarounds that hand-authored ``<a:blipFill>`` XML via
``lxml.etree`` and manually registered relationships on the slide part.

The supported API ``FillFormat.blip_fill(image_file)`` was shipped per
``FEATURES.md`` and resolves #419. This suite is the *verify-close* pass
— it pins the reporter's end-to-end workflow (call ``.blip_fill(...)``
on ``shape.fill``, save, reopen) against silent regression. Each test
exercises the real public API from the caller's perspective and uses no
mocks; shape-level round-trip tests go through
``Presentation.save`` + reopen so the assertion covers the serialized
XML and not just the in-memory proxy.

Sister test ``tests/test_issue_234_blip_fill.py`` already covers the
primary round-trip for an auto-shape plus image-part reuse and the
no-part error path. This file complements #234 by adding coverage for
alternate input forms (BytesIO and ``pathlib.Path``), a second preset
geometry (oval), direct XML assertion on the ``<p:sp>/<p:spPr>``
subtree, and an explicit ``MSO_FILL.PICTURE`` fill-type check.
"""

from __future__ import annotations

import io
from pathlib import Path

from pptx import Presentation
from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches


class DescribeIssue419ShapePictureFill:
    """Verify-close regression suite for picture-fill on shapes (issue #419)."""

    def it_applies_a_picture_fill_to_a_rectangle_autoshape(self):
        # -- the canonical #419 flow: add a rectangle, call blip_fill with a
        # -- filesystem path, confirm the fill is now picture-typed.
        prs, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(2)
        )

        rect.fill.blip_fill("tests/test_files/python-powered.png")

        assert rect.fill.type == MSO_FILL.PICTURE

    def it_applies_a_picture_fill_to_an_oval_autoshape(self):
        # -- #419 is not rectangle-specific; any preset geometry that carries
        # -- a p:spPr should accept blip_fill. Oval is a different prstGeom
        # -- ("ellipse") so it pins that the fill-property rewrite does not
        # -- depend on the surrounding geometry element.
        prs, slide = _fresh_slide()
        oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(3), Inches(2))

        oval.fill.blip_fill("tests/test_files/python-powered.png")

        assert oval.fill.type == MSO_FILL.PICTURE
        # -- the geometry under the shape is untouched by the fill change --
        prstGeom = oval._element.xpath(".//a:prstGeom")[0]
        assert prstGeom.get("prst") == "ellipse"

    def it_reports_fill_type_as_MSO_FILL_PICTURE(self):
        # -- the reporter expects MSO_FILL.PICTURE both before and after a
        # -- save + reopen. Anything else (e.g. MSO_FILL.BACKGROUND,
        # -- AttributeError from a missing mapping) would indicate the
        # -- <a:blipFill> element isn't being recognized round-trip.
        prs, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(2)
        )
        rect.fill.blip_fill("tests/test_files/python-powered.png")

        # -- in-memory --
        assert rect.fill.type == MSO_FILL.PICTURE

        # -- round-trip --
        reloaded = _roundtrip(prs)
        reloaded_shape = reloaded.slides[0].shapes[0]
        assert reloaded_shape.fill.type == MSO_FILL.PICTURE

    def it_survives_a_save_and_reopen_round_trip(self):
        # -- a live #419 user scenario is: apply fill, save, hand the file
        # -- to a downstream consumer (PowerPoint, another library). The
        # -- embedded image bytes and the relationship linking them must
        # -- both survive Presentation.save + Presentation(buf).
        image_path = "tests/test_files/python-powered.png"
        with open(image_path, "rb") as f:
            expected_bytes = f.read()

        prs, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(2)
        )
        rect.fill.blip_fill(image_path)

        reloaded = _roundtrip(prs)
        reloaded_shape = reloaded.slides[0].shapes[0]

        # -- picture fill present and the embedded blob matches on disk --
        assert reloaded_shape.fill.type == MSO_FILL.PICTURE
        blips = reloaded_shape._element.xpath(".//a:blipFill/a:blip")
        assert len(blips) == 1
        rId = blips[0].rEmbed
        assert rId is not None
        image_part = reloaded.slides[0].part.related_part(rId)
        assert image_part.blob == expected_bytes

    def it_accepts_a_BytesIO_stream_as_the_image_file(self):
        # -- #419 commenters asked for file-like input so bytes produced
        # -- in-memory (PIL buffer, a download, a thumbnailer) need not
        # -- hit disk first; blip_fill must accept any readable stream.
        with open("tests/test_files/python-powered.png", "rb") as f:
            image_bytes = f.read()

        _, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(2)
        )

        rect.fill.blip_fill(io.BytesIO(image_bytes))

        assert rect.fill.type == MSO_FILL.PICTURE
        # -- the bytes landed in the image part verbatim --
        blip = rect._element.xpath(".//a:blipFill/a:blip")[0]
        image_part = slide.part.related_part(blip.rEmbed)
        assert image_part.blob == image_bytes

    def it_accepts_a_pathlib_Path_converted_to_str(self):
        # -- the #419 thread includes users whose image paths come from
        # -- ``pathlib.Path`` objects (the common case in modern Python).
        # -- blip_fill's signature is ``str | IO[bytes]``, so the supported
        # -- idiom is ``blip_fill(str(path))`` — pin that it works.
        path = Path("tests/test_files/python-powered.png")
        assert path.exists()  # -- sanity --

        _, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(2)
        )

        rect.fill.blip_fill(str(path))

        assert rect.fill.type == MSO_FILL.PICTURE
        with open(path, "rb") as f:
            expected_bytes = f.read()
        blip = rect._element.xpath(".//a:blipFill/a:blip")[0]
        image_part = slide.part.related_part(blip.rEmbed)
        assert image_part.blob == expected_bytes

    def it_emits_a_blipFill_element_under_the_shape_spPr(self):
        # -- #419's acceptance criterion on the XML side is that the
        # -- shape's ``<p:sp>/<p:spPr>`` acquires an ``<a:blipFill>`` child
        # -- (replacing any prior EG_FillProperties such as a:solidFill
        # -- from the default style). This test asserts that structure
        # -- directly rather than relying on the fill-type accessor, so a
        # -- regression in FillFormat.type doesn't mask the real breakage.
        _, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(2)
        )

        rect.fill.blip_fill("tests/test_files/python-powered.png")

        # -- locate <p:sp>/<p:spPr> and confirm it directly parents a
        # -- single <a:blipFill> that carries an <a:blip r:embed=".."/>
        # -- and an <a:stretch>/<a:fillRect/> child.
        sp = rect._element
        assert sp.tag == qn("p:sp")
        spPrs = sp.findall(qn("p:spPr"))
        assert len(spPrs) == 1
        spPr = spPrs[0]
        blipFills = spPr.findall(qn("a:blipFill"))
        assert len(blipFills) == 1
        blipFill = blipFills[0]
        # -- exactly one <a:blip> with an r:embed attribute --
        blips = blipFill.findall(qn("a:blip"))
        assert len(blips) == 1
        assert blips[0].get(qn("r:embed")) is not None
        # -- default stretch-to-fill present --
        assert blipFill.find(qn("a:stretch") + "/" + qn("a:fillRect")) is not None
        # -- the prior default-style solidFill has been replaced, not duplicated --
        assert spPr.find(qn("a:solidFill")) is None

    def it_survives_round_trip_with_intact_spPr_blipFill_xml(self):
        # -- end-to-end XML check: after save + reopen, the reloaded shape's
        # -- <p:sp>/<p:spPr> still contains exactly one <a:blipFill> with a
        # -- resolvable r:embed. This is the pin against a serializer that
        # -- silently drops <a:blipFill> or reorders children such that the
        # -- relationship fails to resolve.
        prs, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(2)
        )
        rect.fill.blip_fill("tests/test_files/python-powered.png")

        reloaded = _roundtrip(prs)
        reloaded_shape = reloaded.slides[0].shapes[0]

        sp = reloaded_shape._element
        spPr = sp.findall(qn("p:spPr"))[0]
        blipFills = spPr.findall(qn("a:blipFill"))
        assert len(blipFills) == 1
        rEmbed = blipFills[0].findall(qn("a:blip"))[0].get(qn("r:embed"))
        assert rEmbed is not None
        # -- the relationship resolves to a real ImagePart --
        assert reloaded.slides[0].part.related_part(rEmbed) is not None


# -- helpers --------------------------------------------------------------


def _fresh_slide():
    """Return ``(prs, slide)`` — a new ``Presentation`` with one blank slide."""
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    return prs, slide


def _roundtrip(prs):
    """Serialize `prs` to a BytesIO and reopen it. Returns the reloaded Presentation."""
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)
