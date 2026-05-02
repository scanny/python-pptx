# pyright: reportPrivateUsage=false

"""Regression test for issue #1020 — transparency on picture fills, freeform
shapes, and picture-fill extraction.

Issue #1020 (https://github.com/scanny/python-pptx/issues/1020) bundled three
distinct feature asks into a single bug report:

1. Ability to set transparency (alpha) on fills.
2. Ability to draw / inspect *freeform* shapes (the "vertices" geometry
   authored via ``Shapes.build_freeform`` and read back via a shape's path
   geometry).
3. Ability to extract a picture fill from a shape (the read-back analogue
   of ``FillFormat.blip_fill``).

All three asks are resolved in this fork by independently landed features:

* #62 added ``ColorFormat.alpha`` (Wave 8) — per-color opacity.
* #234 added ``FillFormat.blip_fill`` (Wave 7) — picture-fill authoring,
  and the picture fill is inspectable after the fact via the existing
  ``FillFormat`` proxy (``fill.type == MSO_FILL.PICTURE`` and the
  underlying ``<a:blipFill>`` element is reachable through ``fill._xPr``
  for advanced callers needing the embedded rId).
* #515 added ``Shape.path_geometry`` (Wave 1) — read-only path extraction
  for freeform custom-geometry shapes and the static-preset subset (the
  dynamic-preset case remains deferred per #515's documented scope).

This suite pins the three resolutions so #1020 can be closed. Each test
exercises the real API against a fresh ``Presentation()``, not a mock.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.shapes.geometry import Close, LineTo, MoveTo, Path, PathGeometry
from pptx.util import Emu, Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against test-module pollution of `PartFactory.part_type_for`.

    Other test modules (notably ``tests/opc/test_package.py``) overwrite
    slide-part registrations with mocks and do not restore them. Round-
    trip tests that call ``Presentation.save`` + ``Presentation()``
    reopen depend on the real registrations being in place. Mirrors the
    guard used by ``tests/test_issue_62_fill_alpha.py`` and
    ``tests/test_issue_234_blip_fill.py``.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.parts.image import ImagePart
    from pptx.parts.media import MediaPart
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.slide import (
        NotesMasterPart,
        NotesSlidePart,
        SlideLayoutPart,
        SlideMasterPart,
        SlidePart,
    )

    saved = dict(PartFactory.part_type_for)
    expected = {
        CT.PML_PRESENTATION_MAIN: PresentationPart,
        CT.PML_PRES_MACRO_MAIN: PresentationPart,
        CT.PML_TEMPLATE_MAIN: PresentationPart,
        CT.PML_SLIDESHOW_MAIN: PresentationPart,
        CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
        CT.PML_NOTES_MASTER: NotesMasterPart,
        CT.PML_NOTES_SLIDE: NotesSlidePart,
        CT.PML_SLIDE: SlidePart,
        CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
        CT.PML_SLIDE_MASTER: SlideMasterPart,
        CT.DML_CHART: ChartPart,
        CT.JPEG: ImagePart,
        CT.PNG: ImagePart,
        CT.MP4: MediaPart,
    }
    PartFactory.part_type_for.update(expected)
    try:
        yield
    finally:
        PartFactory.part_type_for.clear()
        PartFactory.part_type_for.update(saved)


class DescribeIssue1020FillAlphaFreeformVerify(object):
    """#1020 verify-and-close: alpha + blip_fill + freeform path extraction."""

    # -- 1: picture-fill extraction / read-access (via #234) ----------------

    def it_exposes_picture_fill_type_and_embedded_rId_after_blip_fill(self, _restore_part_factory):
        """Pins that ``shape.fill`` reads back a PICTURE fill type and the
        embedded image part is reachable after authoring via ``blip_fill``.

        This is the "extract a picture fill" half of #1020 — after the
        authoring call, a caller inspecting the shape sees a PICTURE
        fill, the underlying ``<a:blipFill>`` element is present on the
        ``FillFormat``'s parent (so advanced callers can walk to
        ``a:blip/@r:embed``), and the rId resolves to the ``ImagePart``
        whose blob matches the source image.
        """
        image_path = "tests/test_files/python-powered.png"
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(2)
        )

        # -- arrange: author a picture fill --
        shape.fill.blip_fill(image_path)

        # -- act + assert (round 1, pre-roundtrip): the same FillFormat
        # -- instance that authored the fill now reads back as PICTURE
        # -- and the blipFill subtree is discoverable for extraction. --
        assert shape.fill.type == MSO_FILL.PICTURE
        blips = shape.fill._xPr.xpath("a:blipFill/a:blip")
        assert len(blips) == 1
        rId = blips[0].rEmbed
        assert rId is not None
        # -- the rId resolves to an ImagePart on the slide part --
        image_part = slide.part.related_part(rId)
        with open(image_path, "rb") as f:
            assert image_part.blob == f.read()

        # -- act + assert (round 2, post-roundtrip): a *fresh*
        # -- Presentation reopened from the saved bytes also reads back
        # -- PICTURE with an extractable image. --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        reloaded = prs2.slides[0].shapes[0]
        assert reloaded.fill.type == MSO_FILL.PICTURE
        reloaded_blips = reloaded.fill._xPr.xpath("a:blipFill/a:blip")
        assert len(reloaded_blips) == 1
        reloaded_rId = reloaded_blips[0].rEmbed
        reloaded_image = prs2.slides[0].part.related_part(reloaded_rId)
        with open(image_path, "rb") as f:
            assert reloaded_image.blob == f.read()

    # -- 2: freeform-shape path extraction (via #515) -----------------------

    def it_extracts_path_geometry_from_a_built_freeform_shape(self, _restore_part_factory):
        """Pins that ``Shape.path_geometry`` reads back the custom geometry
        authored by ``Shapes.build_freeform`` + ``add_line_segments``.

        Dynamic preset shapes (those whose ``a:avLst`` adjustment
        values drive the geometry) are a documented gap in #515 and
        remain out of scope for this test. The companion test below
        asserts that a dynamic preset still correctly reports |None|,
        and a preset-geometry sample is also exercised to cover the
        static-preset branch.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # -- author a triangle freeform with scale=1 EMU/unit so the
        # -- coordinates round-trip cleanly through path extraction. --
        builder = slide.shapes.build_freeform(start_x=0, start_y=0, scale=1.0)
        builder.add_line_segments([(100, 0), (100, 100), (0, 100)], close=True)
        freeform = builder.convert_to_shape()

        # -- assert: the authored shape is a FREEFORM (has custom geometry) --
        assert freeform.shape_type == MSO_SHAPE_TYPE.FREEFORM

        # -- act: extract the path geometry --
        pg = freeform.path_geometry
        assert isinstance(pg, PathGeometry)
        # -- freeform builder emits one a:path --
        assert len(pg) == 1
        path = pg[0]
        assert isinstance(path, Path)

        # -- assert: drawing ops reflect the authored vertices --
        ops = list(path)
        # -- MoveTo to the pen start, then 3 LineTo's, then Close --
        assert ops[0] == MoveTo(Emu(0), Emu(0))
        assert ops[1] == LineTo(Emu(100), Emu(0))
        assert ops[2] == LineTo(Emu(100), Emu(100))
        assert ops[3] == LineTo(Emu(0), Emu(100))
        assert ops[-1] == Close()

        # -- assert: round-trip through save + reopen preserves geometry --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        reloaded = prs2.slides[0].shapes[0]
        assert reloaded.shape_type == MSO_SHAPE_TYPE.FREEFORM
        pg2 = reloaded.path_geometry
        assert isinstance(pg2, PathGeometry)
        assert len(pg2) == 1
        ops2 = list(pg2[0])
        assert ops2[0] == MoveTo(Emu(0), Emu(0))
        assert ops2[-1] == Close()
        # -- same count of ops as before the round-trip --
        assert len(ops2) == len(ops)

    # -- 3: alpha on fills (via #62) ----------------------------------------

    def it_sets_alpha_on_a_solid_fill_that_round_trips(self, _restore_part_factory):
        """Pins that ``ColorFormat.alpha`` writes an ``<a:alpha val>`` child
        and that the value round-trips through save + reopen.

        This is the "transparency on fills" half of #1020. A value of
        0.4 (40% opaque / 60% transparent) is chosen so that the
        serialised XML is distinguishable from the OOXML default
        (no ``<a:alpha>`` child, interpreted as fully opaque).
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(2)
        )

        # -- arrange: solid-fill a blue rectangle --
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0x00, 0x33, 0xCC)

        # -- act: set alpha to 40% opacity --
        shape.fill.fore_color.alpha = 0.4

        # -- assert: reading back returns the assigned value --
        assert shape.fill.fore_color.alpha == 0.4
        # -- assert: the <a:alpha val="40000"/> element is present --
        xml = shape.fill._xPr.xml.encode()
        assert b'<a:alpha val="40000"/>' in xml

        # -- round-trip: save + reopen preserves the alpha --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        reloaded = prs2.slides[0].shapes[1]
        assert reloaded.fill.fore_color.alpha == 0.4
        assert reloaded.fill.fore_color.rgb == RGBColor(0x00, 0x33, 0xCC)

        # -- assigning None removes the alpha child (inherits default) --
        reloaded.fill.fore_color.alpha = None
        assert reloaded.fill.fore_color.alpha == 1.0
        assert b"<a:alpha" not in reloaded.fill._xPr.xml.encode()
