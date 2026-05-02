# pyright: reportPrivateUsage=false

"""Regression test for issue #434 — ``InvalidXmlError`` on image read.

Issue #434 (https://github.com/scanny/python-pptx/issues/434) reports that
opening a ``.pptx`` whose ``p:pic`` element lacks a ``p:blipFill`` child
raises :class:`pptx.exc.InvalidXmlError` at shape-factory construction — or
the first time any crop / image property is accessed. A ``p:blipFill`` can
legitimately go missing when another tool strips image data out of a
presentation (e.g. for size / PII reasons), or when an intermediate
conversion truncates the element tree. Such a shape is malformed, but
python-pptx historically raised instead of giving callers a way to inspect
it, skip it, or reflow around it.

The fix relaxes ``CT_Picture.blipFill`` from ``OneAndOnlyOne`` to
``ZeroOrOne`` and teaches every ``blipFill``-dependent accessor
(``blip_rId``, ``_srcRect_x``, ``Picture.image``) to handle the missing-child
case. Reading ``.image`` now returns |None| instead of raising, and the
crop accessors return 0.0, matching the "no srcRect" behaviour.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.shapes.picture import Picture
from pptx.util import Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against test-module pollution of `PartFactory.part_type_for`.

    See the identical fixture in ``tests/test_issue_806_linked_picture.py``.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.comments import CommentAuthorsPart, CommentsPart
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
        CT.PML_COMMENTS: CommentsPart,
        CT.PML_COMMENT_AUTHORS: CommentAuthorsPart,
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


def _pic_without_blipFill():
    """Return a ``CT_Picture`` element missing its ``p:blipFill`` child.

    All other required sub-elements (``p:nvPicPr``, ``p:spPr``) are present
    so the element parses. This is the shape of the element tree in the
    malformed packages reported on issue #434.
    """
    xml = (
        "<p:pic %s>\n"
        "  <p:nvPicPr>\n"
        '    <p:cNvPr id="4" name="Pic 1" descr=""/>\n'
        '    <p:cNvPicPr>\n'
        '      <a:picLocks noChangeAspect="1"/>\n'
        "    </p:cNvPicPr>\n"
        "    <p:nvPr/>\n"
        "  </p:nvPicPr>\n"
        "  <p:spPr>\n"
        "    <a:xfrm>\n"
        '      <a:off x="914400" y="914400"/>\n'
        '      <a:ext cx="2743200" cy="2057400"/>\n'
        "    </a:xfrm>\n"
        '    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>\n'
        "  </p:spPr>\n"
        "</p:pic>" % nsdecls("p", "a", "r")
    )
    return parse_xml(xml)


class DescribeIssue434PictureMissingBlipFill:
    """A ``p:pic`` without ``p:blipFill`` no longer raises on access."""

    def it_constructs_a_Picture_proxy_without_raising(self):
        """Shape-factory construction and property access must succeed."""
        pic = _pic_without_blipFill()

        shape = Picture(pic, None)

        assert isinstance(shape, Picture)

    def it_returns_None_for_image_instead_of_raising(self):
        """``Picture.image`` is the primary access point flagged by #434.

        It used to propagate ``InvalidXmlError`` via ``blip_rId``; now it
        returns |None| so callers can branch cleanly.
        """
        pic = _pic_without_blipFill()
        shape = Picture(pic, None)

        assert shape.image is None

    def it_reports_zero_cropping_when_blipFill_is_absent(self):
        """Each crop accessor must return 0.0 (the "no srcRect" default)."""
        pic = _pic_without_blipFill()
        shape = Picture(pic, None)

        assert shape.crop_left == 0.0
        assert shape.crop_top == 0.0
        assert shape.crop_right == 0.0
        assert shape.crop_bottom == 0.0

    def it_reports_None_for_blip_rId_when_blipFill_is_absent(self):
        """The oxml property that actually touches ``p:blipFill`` must be
        safe. ``CT_Picture.blip_rId`` is shared by ``Picture.image``,
        ``Movie.poster_frame``, and ``Movie._media_rIds``.
        """
        pic = _pic_without_blipFill()

        assert pic.blip_rId is None

    def it_raises_ValueError_when_replace_image_has_no_blipFill(self):
        """``replace_image`` still rejects malformed shapes — the shape has
        no "current" image bytes to swap, so returning silently would hide
        a caller bug. The error type is ``ValueError`` (not
        ``InvalidXmlError``) — a clean, documented API contract.
        """
        pic = _pic_without_blipFill()
        shape = Picture(pic, None)

        with pytest.raises(ValueError, match="no embedded image to replace"):
            shape.replace_image(io.BytesIO(b"does not matter"))

    def it_round_trips_a_pic_without_blipFill_through_save_and_reopen(
        self, _restore_part_factory
    ):
        """End-to-end: a presentation carrying a blipFill-less ``p:pic``
        opens, its picture is reachable via ``slide.shapes[...]``, and
        ``Picture.image`` returns |None| — all without raising.
        """
        # -- build a normal one-picture presentation, then surgically
        #    remove the p:blipFill on disk before reopening --
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.add_picture(
            "tests/test_files/python-powered.png",
            Inches(1),
            Inches(1),
            Inches(2),
            Inches(2),
        )
        # -- strip the blipFill element from the picture before save --
        pic_elm = slide.shapes[-1]._element
        blipFill = pic_elm.blipFill
        assert blipFill is not None, "test precondition: pic starts with a blipFill"
        pic_elm.remove(blipFill)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        # -- reopen: factory must not raise, image must be None --
        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]
        picture2 = slide2.shapes[-1]

        assert isinstance(picture2, Picture)
        assert picture2.image is None
        # -- crops still report their default 0.0 --
        assert picture2.crop_left == 0.0
