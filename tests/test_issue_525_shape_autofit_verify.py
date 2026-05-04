# pyright: reportPrivateUsage=false

"""Regression test for issue #525 — adjust textbox height to fit text.

Issue #525 (https://github.com/scanny/python-pptx/issues/525) asked for a way
to have a textbox auto-grow vertically to fit its text, instead of the text
overflowing the shape's fixed bounds.

The capability is a native PowerPoint feature expressed via the text body's
``a:bodyPr`` auto-fit choice:

* ``a:spAutoFit`` — shape auto-resizes to fit the text (the issue-#525 ask).
* ``a:normAutofit`` — text shrinks to fit the shape.
* ``a:noAutofit`` — no auto-fit (text overflows silently).

python-pptx exposes these through :attr:`TextFrame.auto_size` as
:class:`MSO_AUTO_SIZE` members: ``SHAPE_TO_FIT_TEXT``, ``TEXT_TO_FIT_SHAPE``,
and ``NONE`` respectively. Setting ``auto_size = None`` removes any auto-fit
choice so the setting is inherited from the layout/master/theme.

python-pptx cannot *compute* the new shape bounds — that requires per-font
metrics for arbitrary user fonts, which PowerPoint applies at render time.
What the library *does* do, and this suite pins, is emit the
``a:spAutoFit`` element so that PowerPoint performs the resize the first
time the file is opened.

This suite verifies the four round-trip behaviours of ``TextFrame.auto_size``
on a plain autoshape so #525 can be closed as already-resolved.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against test-module pollution of `PartFactory.part_type_for`.

    Other test modules (notably ``tests/opc/test_package.py``) overwrite
    slide-part registrations with mocks and do not restore them. Round-
    trip tests that call ``Presentation.save`` + ``Presentation()``
    reopen depend on the real registrations being in place.
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


def _make_textbox_shape():
    """Return a fresh (Presentation, shape) pair with a plain rectangle shape.

    The rectangle carries a default ``a:bodyPr`` (no auto-fit choice) so that
    every test starts from the same well-defined baseline.
    """
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
    shape.text_frame.text = "Some text that might overflow"
    return prs, shape


def _reopen(prs):
    """Round-trip ``prs`` through in-memory save + Presentation()."""
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


class DescribeIssue525ShapeAutofitVerify(object):
    """#525 verify-and-close: TextFrame.auto_size writes the correct a:bodyPr child."""

    # -- 1: NONE emits neither spAutoFit nor normAutofit --------------------

    def it_emits_noAutofit_for_MSO_AUTO_SIZE_NONE(self, _restore_part_factory):
        """``auto_size = MSO_AUTO_SIZE.NONE`` writes ``a:noAutofit`` and no
        other auto-fit choice sibling."""
        prs, shape = _make_textbox_shape()
        tf = shape.text_frame

        tf.auto_size = MSO_AUTO_SIZE.NONE

        assert tf.auto_size is MSO_AUTO_SIZE.NONE
        xml = tf._txBody.xml
        assert "<a:noAutofit" in xml
        assert "<a:spAutoFit" not in xml
        assert "<a:normAutofit" not in xml

        # -- round-trip --
        prs2 = _reopen(prs)
        tf2 = prs2.slides[0].shapes[1].text_frame
        assert tf2.auto_size is MSO_AUTO_SIZE.NONE
        xml2 = tf2._txBody.xml
        assert "<a:noAutofit" in xml2
        assert "<a:spAutoFit" not in xml2
        assert "<a:normAutofit" not in xml2

    # -- 2: SHAPE_TO_FIT_TEXT emits a:spAutoFit (the #525 ask) --------------

    def it_emits_spAutoFit_for_MSO_AUTO_SIZE_SHAPE_TO_FIT_TEXT(self, _restore_part_factory):
        """``auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`` writes
        ``a:spAutoFit`` so PowerPoint auto-grows the shape to fit its text
        when the file is opened. This is the direct answer to #525."""
        prs, shape = _make_textbox_shape()
        tf = shape.text_frame

        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

        assert tf.auto_size is MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        xml = tf._txBody.xml
        assert "<a:spAutoFit" in xml
        assert "<a:normAutofit" not in xml
        assert "<a:noAutofit" not in xml

        # -- round-trip --
        prs2 = _reopen(prs)
        tf2 = prs2.slides[0].shapes[1].text_frame
        assert tf2.auto_size is MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        xml2 = tf2._txBody.xml
        assert "<a:spAutoFit" in xml2
        assert "<a:normAutofit" not in xml2
        assert "<a:noAutofit" not in xml2

    # -- 3: TEXT_TO_FIT_SHAPE emits a:normAutofit ---------------------------

    def it_emits_normAutofit_for_MSO_AUTO_SIZE_TEXT_TO_FIT_SHAPE(self, _restore_part_factory):
        """``auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`` writes
        ``a:normAutofit`` so PowerPoint shrinks the text to fit the shape."""
        prs, shape = _make_textbox_shape()
        tf = shape.text_frame

        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        assert tf.auto_size is MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        xml = tf._txBody.xml
        assert "<a:normAutofit" in xml
        assert "<a:spAutoFit" not in xml
        assert "<a:noAutofit" not in xml

        # -- round-trip --
        prs2 = _reopen(prs)
        tf2 = prs2.slides[0].shapes[1].text_frame
        assert tf2.auto_size is MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        xml2 = tf2._txBody.xml
        assert "<a:normAutofit" in xml2
        assert "<a:spAutoFit" not in xml2
        assert "<a:noAutofit" not in xml2

    # -- 4: assigning None removes every auto-fit choice --------------------

    def it_removes_all_autofit_choices_when_auto_size_is_set_to_None(self, _restore_part_factory):
        """Assigning ``auto_size = None`` strips whichever auto-fit choice is
        currently present so the setting is inherited from the layout/master
        /theme."""
        prs, shape = _make_textbox_shape()
        tf = shape.text_frame

        # -- arrange: start with spAutoFit written in --
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        assert "<a:spAutoFit" in tf._txBody.xml

        # -- act: remove the auto-fit choice --
        tf.auto_size = None

        # -- assert: no auto-fit element of any kind remains --
        assert tf.auto_size is None
        xml = tf._txBody.xml
        assert "<a:spAutoFit" not in xml
        assert "<a:normAutofit" not in xml
        assert "<a:noAutofit" not in xml

        # -- round-trip preserves the None state --
        prs2 = _reopen(prs)
        tf2 = prs2.slides[0].shapes[1].text_frame
        assert tf2.auto_size is None
        xml2 = tf2._txBody.xml
        assert "<a:spAutoFit" not in xml2
        assert "<a:normAutofit" not in xml2
        assert "<a:noAutofit" not in xml2

    # -- 5: switching between choices replaces the prior sibling ------------

    def it_replaces_the_existing_autofit_choice_when_reassigned(self, _restore_part_factory):
        """Successive assignments must swap in a single auto-fit child — not
        accumulate choice siblings (which would violate the
        ``EG_TextAutofit`` XSD choice group)."""
        prs, shape = _make_textbox_shape()
        tf = shape.text_frame

        for value, expected_child in (
            (MSO_AUTO_SIZE.NONE, "<a:noAutofit"),
            (MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT, "<a:spAutoFit"),
            (MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE, "<a:normAutofit"),
            (MSO_AUTO_SIZE.NONE, "<a:noAutofit"),
        ):
            tf.auto_size = value
            xml = tf._txBody.xml
            assert expected_child in xml
            # -- count total auto-fit children; must be exactly one --
            total = (
                xml.count("<a:noAutofit") + xml.count("<a:normAutofit") + xml.count("<a:spAutoFit")
            )
            assert total == 1, f"expected exactly one auto-fit choice child, found {total}"
