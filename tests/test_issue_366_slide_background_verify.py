# pyright: reportPrivateUsage=false

"""Regression test for issue #366 — ``Slide.background``.

Issue #366 (https://github.com/scanny/python-pptx/issues/366) asked for
a way to set a slide's background — solid color, gradient, or picture —
from python-pptx, and to revert back to master-inherited background
afterwards. The feature is implemented by the ``_Background`` proxy
returned from :attr:`.Slide.background`, whose ``.fill`` attribute
exposes the standard :class:`.FillFormat` surface (``solid()``,
``gradient()``, ``blip_fill(image_file)``) — exactly the same API that
already powers shape / table-cell / data-label backgrounds.

The complementary read-side hook is :attr:`.Slide.follow_master_background`
— ``True`` when the slide has no ``p:bg`` override, ``False`` when an
explicit background has been applied. Removing the explicit ``p:bg``
restores inheritance from the slide master / layout.

This suite pins the four scenarios from the reporter's perspective:

1. A fresh slide inherits its background from the master — no ``p:bg``
   element is written, and :attr:`.Slide.follow_master_background` is
   ``True``.
2. Authoring a *solid* background color via
   ``slide.background.fill.solid()`` +
   ``slide.background.fill.fore_color.rgb = RGBColor(...)`` round-trips
   cleanly through ``Presentation.save`` + reopen.
3. Authoring a *gradient* background via
   ``slide.background.fill.gradient()`` round-trips — ``fill.type`` is
   ``MSO_FILL.GRADIENT`` on the reloaded deck.
4. Authoring a *picture* background via
   ``slide.background.fill.blip_fill(image_file)`` round-trips — the
   embedded image part survives and ``fill.type`` is
   ``MSO_FILL.PICTURE`` on the reloaded deck.
5. Clearing the ``p:bg`` element reverts the slide to master-inherited
   background — :attr:`.Slide.follow_master_background` flips back to
   ``True`` and no ``p:bg`` child remains under ``p:cSld``.

Any regression that drops the ``p:bg`` round-trip, fails to register
the embedded image part, or stops removing the override when cleared
will reproduce the #366 symptom and be caught here.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL
from pptx.oxml.ns import qn


@pytest.fixture
def _restore_part_factory():
    """Guard against test-module pollution of ``PartFactory.part_type_for``.

    Mirrors the guard used by ``tests/test_issue_234_blip_fill.py`` and
    friends — ``Presentation(stream)`` reopen dispatches through
    ``PartFactory.part_type_for``, which earlier test modules overwrite
    with mocks and do not restore.
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


class DescribeIssue366SlideBackgroundVerify:
    """#366 slide-background verify-and-close.

    Pins the solid / gradient / picture authoring flows and the
    master-inheritance read path so the #366 report can be closed.
    """

    # -- 1. Default: master-inherited background --------------------------

    def it_reports_master_inherited_background_on_a_fresh_slide(self):
        """A freshly-added slide has no ``p:bg`` override.

        ``Slide.follow_master_background`` returns ``True`` and the
        slide's ``p:cSld`` carries no ``p:bg`` child. This is the
        starting state the #366 reporter expected to flip *away from*
        when applying a custom background.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        assert bool(slide.follow_master_background) is True
        # -- no p:bg child written under p:cSld --
        assert slide._element.cSld.bg is None
        assert slide._element.cSld.find(qn("p:bg")) is None

    # -- 2. Solid color background round-trip -----------------------------

    def it_round_trips_a_solid_color_background(self, _restore_part_factory):
        """Author a solid-fill background and reopen.

        The classic #366 ask: set the slide background to a specific
        RGB color. Authoring writes a ``p:bg/p:bgPr/a:solidFill`` subtree
        and flips ``follow_master_background`` to ``False``. On reopen
        the fill type is still ``MSO_FILL.SOLID`` with the original RGB.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        # -- act: apply a solid red background --
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)

        # -- assert: in-memory state reflects the override --
        assert bool(slide.follow_master_background) is False
        assert slide.background.fill.type == MSO_FILL.SOLID
        assert slide.background.fill.fore_color.rgb == RGBColor(0xFF, 0x00, 0x00)

        # -- round-trip --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]

        # -- assert: reloaded slide keeps the solid background --
        assert bool(slide2.follow_master_background) is False
        assert slide2.background.fill.type == MSO_FILL.SOLID
        assert slide2.background.fill.fore_color.rgb == RGBColor(0xFF, 0x00, 0x00)
        # -- and the p:bg subtree is still present --
        assert slide2._element.cSld.bg is not None

    # -- 3. Gradient background round-trip --------------------------------

    def it_round_trips_a_gradient_background(self, _restore_part_factory):
        """Author a gradient background and reopen.

        ``FillFormat.gradient()`` stamps the default linear two-stop
        accent1 gradient (the same default PowerPoint's "White" template
        uses). The fill type survives save + reopen.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        # -- act: apply the default gradient --
        slide.background.fill.gradient()

        # -- assert: in-memory state --
        assert bool(slide.follow_master_background) is False
        assert slide.background.fill.type == MSO_FILL.GRADIENT
        # -- default gradient has at least two stops --
        assert len(slide.background.fill.gradient_stops) >= 2

        # -- round-trip --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]

        assert bool(slide2.follow_master_background) is False
        assert slide2.background.fill.type == MSO_FILL.GRADIENT
        assert len(slide2.background.fill.gradient_stops) >= 2

    # -- 4. Picture background round-trip ---------------------------------

    def it_round_trips_a_picture_background(self, _restore_part_factory):
        """Author a picture background via ``blip_fill`` and reopen.

        The picture bytes must be embedded on the slide part, the fill
        type must be ``MSO_FILL.PICTURE``, and on reopen the same image
        part must still resolve from the ``a:blip/@r:embed`` rId under
        ``p:bg/p:bgPr``.
        """
        image_path = "tests/test_files/python-powered.png"
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        # -- act: apply a picture background --
        slide.background.fill.blip_fill(image_path)

        # -- assert: in-memory state --
        assert bool(slide.follow_master_background) is False
        assert slide.background.fill.type == MSO_FILL.PICTURE
        # -- a:blipFill lives under p:bg/p:bgPr --
        blips = slide._element.xpath(".//p:cSld/p:bg/p:bgPr/a:blipFill/a:blip")
        assert len(blips) == 1
        rId = blips[0].rEmbed
        assert rId is not None
        assert slide.part.related_part(rId) is not None

        # -- round-trip --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]

        assert bool(slide2.follow_master_background) is False
        assert slide2.background.fill.type == MSO_FILL.PICTURE
        reloaded_blips = slide2._element.xpath(".//p:cSld/p:bg/p:bgPr/a:blipFill/a:blip")
        assert len(reloaded_blips) == 1
        reloaded_rId = reloaded_blips[0].rEmbed
        reloaded_image_part = slide2.part.related_part(reloaded_rId)
        # -- bytes match the source image --
        with open(image_path, "rb") as f:
            expected_bytes = f.read()
        assert reloaded_image_part.blob == expected_bytes

    # -- 5. Revert to master-inherited background -------------------------

    def it_reverts_to_master_inherited_background_when_p_bg_is_removed(self):
        """Removing ``p:bg`` restores master inheritance.

        After authoring a solid-color background the ``p:cSld`` carries
        a ``p:bg`` child. Removing that child — the write-side of the
        :attr:`.Slide.follow_master_background` contract documented in
        ``docs/dev/analysis/sld-background.rst`` — flips
        ``follow_master_background`` back to ``True`` and leaves no
        ``p:bg`` under ``p:cSld``.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        # -- arrange: apply a custom background so p:bg exists --
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0x33, 0x66, 0x99)
        assert bool(slide.follow_master_background) is False
        assert slide._element.cSld.bg is not None

        # -- act: revert to master inheritance by removing p:bg --
        slide._element.cSld._remove_bg()

        # -- assert: master inheritance is restored --
        assert bool(slide.follow_master_background) is True
        assert slide._element.cSld.bg is None
        assert slide._element.cSld.find(qn("p:bg")) is None

    def it_round_trips_the_revert_through_save_and_reload(self, _restore_part_factory):
        """The reverted state survives save + reopen.

        Authoring, reverting, saving and reopening must produce a
        deck whose slide still has no ``p:bg`` override — any regression
        that re-emitted a stale ``p:bg`` from cached proxy state would be
        caught here.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0x00, 0x80, 0x00)
        # -- revert before saving --
        slide._element.cSld._remove_bg()

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]

        assert bool(slide2.follow_master_background) is True
        assert slide2._element.cSld.bg is None
