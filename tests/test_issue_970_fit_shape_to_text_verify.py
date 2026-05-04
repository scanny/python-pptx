# pyright: reportPrivateUsage=false

"""Regression test for issue #970 — "Fit shape to text?"

Issue #970 (https://github.com/scanny/python-pptx/issues/970) asked whether
python-pptx exposes PowerPoint's "Resize shape to fit text" behaviour — the
option in the "Format Shape -> Text Options -> Text Box" pane that tells
PowerPoint to grow (or shrink) the shape's bounding box so the text fits
inside with no overflow. In OOXML the setting is a one-shot choice element
inside ``a:bodyPr``:

* ``a:spAutoFit`` — shape auto-resizes to fit the text (the #970 ask).
* ``a:normAutofit`` — text shrinks to fit the shape.
* ``a:noAutofit`` — no auto-fit (text overflows silently).

python-pptx has always exposed these through :attr:`TextFrame.auto_size` as
:class:`MSO_AUTO_SIZE` members. Assigning
``MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`` writes ``a:spAutoFit`` into the
text-body's ``a:bodyPr``, and PowerPoint performs the resize the first time
the file is opened. Assigning ``None`` removes the auto-fit choice
altogether so the setting falls back to the layout/master/theme default.

#970 is therefore a duplicate of / overlaps #525 ("adjust textbox height to
fit text"). The capability was already present when #970 was filed — the
reporter was looking for the public-API name. The fork-era #525 verify
("verify-and-close" in ``tests/test_issue_525_shape_autofit_verify.py``)
pins the full matrix of auto-fit choices across save+reopen; this suite
focuses specifically on the #970 wording ("fit shape to text") and the five
guarantees #970's reporter is after:

  1. Setting ``auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`` on a plain
     textbox succeeds and :attr:`TextFrame.auto_size` reads the setting
     back.
  2. ``a:spAutoFit`` is present in the serialized XML (no ``a:noAutofit``
     or ``a:normAutofit`` sibling).
  3. The setting survives ``Presentation.save`` + reopen (save to a
     ``BytesIO`` and re-open is the #970 scenario — the deck is produced
     by python-pptx and then opened in PowerPoint).
  4. Setting ``auto_size = MSO_AUTO_SIZE.NONE`` (and assigning ``None``)
     both remove ``a:spAutoFit`` from the XML, so a caller can clear a
     previously-set "resize shape to fit text" flag.
  5. A brief cross-reference to #525 and #715 — the two closely adjacent
     issues in the python-pptx tracker that also exercise the
     ``a:bodyPr`` auto-fit-choice machinery. #525 is the textbox
     auto-grow ask this issue duplicates; #715 is the placeholder
     text-shrink ask that lands on ``a:normAutofit/@fontScale`` and
     ``@lnSpcReduction`` (the read/write :attr:`TextFrame.font_scale`
     and :attr:`TextFrame.line_space_reduction` properties).

The library cannot *compute* the new shape bounds — that requires
per-font metrics for arbitrary user fonts, which PowerPoint applies at
render time. What the library *does* do, and this suite pins, is emit the
``a:spAutoFit`` element so that PowerPoint performs the resize the first
time the file is opened.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.text.text import TextFrame
from pptx.util import Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against test-module pollution of ``PartFactory.part_type_for``.

    Mirrors the guard used by ``tests/test_issue_525_shape_autofit_verify.py``
    and friends. Other test modules (notably ``tests/opc/test_package.py``)
    overwrite slide-part registrations with mocks and do not restore them;
    round-trip tests that call ``Presentation.save`` + ``Presentation()``
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


def _make_textbox(text: str = "Some text that might overflow"):
    """Return ``(prs, textbox)`` with a freshly-authored textbox on slide 0.

    The reporter's #970 scenario is a textbox authored via
    :meth:`SlideShapes.add_textbox` — not an autoshape — because the
    "Format Shape -> Text Options -> Text Box -> Resize shape to fit text"
    toggle is primarily a textbox feature in PowerPoint's UI. A textbox
    authored via ``add_textbox`` carries a default ``a:bodyPr`` with no
    auto-fit choice child, so every test starts from a well-defined
    baseline.
    """
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
    textbox.text_frame.text = text
    return prs, textbox


def _reopen(prs):
    """Round-trip ``prs`` through in-memory save + Presentation()."""
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


class DescribeIssue970FitShapeToTextVerify(object):
    """#970 verify-and-close: "Fit shape to text?" resolved by #525/#715.

    Pins the five guarantees a #970-class caller needs from
    :attr:`TextFrame.auto_size` when they reach for
    ``MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`` on a textbox.
    """

    # -- 1: setting auto_size on a textbox -----------------------------------

    def it_accepts_SHAPE_TO_FIT_TEXT_on_a_textbox_text_frame(self):
        """Scenario 1: the public-API name #970's reporter is after.

        Assigning ``MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`` to
        :attr:`TextFrame.auto_size` on a plain textbox must succeed and
        round-trip through the getter — this is the direct answer to
        "fit shape to text?" on the python-pptx tracker.
        """
        _, textbox = _make_textbox()
        tf = textbox.text_frame

        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

        # -- getter reports the setter's value --
        assert tf.auto_size is MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    # -- 2: a:spAutoFit lands in the serialized XML --------------------------

    def it_emits_spAutoFit_into_the_bodyPr_element(self):
        """Scenario 2: the setting lands on the wire as ``a:spAutoFit``.

        PowerPoint reads the auto-fit choice from the text-body's
        ``a:bodyPr`` on first open; if ``a:spAutoFit`` is missing (or a
        sibling auto-fit choice is also present) the shape will not
        auto-grow. Pin the serialized XML so a regression to the
        ``xmlchemy`` descriptor layer cannot silently drop the element.
        """
        _, textbox = _make_textbox()
        tf = textbox.text_frame

        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

        xml = tf._txBody.xml
        assert "<a:spAutoFit" in xml
        # -- and no other auto-fit-choice sibling (ST_TextAutoFit is a
        #    one-of choice group in dml-text.xsd) --
        assert "<a:normAutofit" not in xml
        assert "<a:noAutofit" not in xml

    # -- 3: round-trip save + reopen -----------------------------------------

    def it_round_trips_SHAPE_TO_FIT_TEXT_through_save_and_reopen(self, _restore_part_factory):
        """Scenario 3: save to ``.pptx`` and reopen preserves the setting.

        The #970 reporter's scenario is a deck produced by python-pptx and
        then opened in PowerPoint — i.e., the library's job is to produce
        a file whose ``a:bodyPr`` still carries ``a:spAutoFit`` after the
        OPC serialize → zip → unzip → parse round-trip. Pin that end to
        end by saving to a ``BytesIO`` and reopening via
        :class:`Presentation`.
        """
        prs, textbox = _make_textbox()
        tf = textbox.text_frame

        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

        prs2 = _reopen(prs)
        # -- find the reopened textbox; add_textbox placed it after any
        #    layout-cloned shapes, so locate it by its text. --
        tf2 = None
        for shape in prs2.slides[0].shapes:
            if shape.has_text_frame and shape.text_frame.text.startswith("Some text"):
                tf2 = shape.text_frame
                break
        assert tf2 is not None, "reopened textbox not found"

        assert tf2.auto_size is MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        xml2 = tf2._txBody.xml
        assert "<a:spAutoFit" in xml2
        assert "<a:normAutofit" not in xml2
        assert "<a:noAutofit" not in xml2

    # -- 4: clearing the setting ---------------------------------------------

    def it_removes_spAutoFit_when_auto_size_is_set_to_NONE(self):
        """Scenario 4a: ``MSO_AUTO_SIZE.NONE`` swaps ``spAutoFit`` out.

        The ``EG_TextAutofit`` choice group in ``dml-text.xsd`` allows at
        most one auto-fit element on ``a:bodyPr``; assigning
        ``MSO_AUTO_SIZE.NONE`` must replace the existing ``a:spAutoFit``
        with ``a:noAutofit`` — not accumulate siblings. This is the
        "how do I turn it back off?" half of the #970 ask.
        """
        _, textbox = _make_textbox()
        tf = textbox.text_frame

        # -- arrange: start with spAutoFit written in --
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        assert "<a:spAutoFit" in tf._txBody.xml

        # -- act: flip to NONE --
        tf.auto_size = MSO_AUTO_SIZE.NONE

        # -- assert: spAutoFit gone, noAutofit present, no duplicates --
        assert tf.auto_size is MSO_AUTO_SIZE.NONE
        xml = tf._txBody.xml
        assert "<a:spAutoFit" not in xml
        assert "<a:normAutofit" not in xml
        assert "<a:noAutofit" in xml
        # -- exactly one auto-fit-choice child --
        total = xml.count("<a:noAutofit") + xml.count("<a:normAutofit") + xml.count("<a:spAutoFit")
        assert total == 1

    def it_removes_spAutoFit_when_auto_size_is_set_to_None(self):
        """Scenario 4b: assigning ``None`` strips the auto-fit choice.

        ``None`` is distinct from ``MSO_AUTO_SIZE.NONE`` — the latter
        writes ``a:noAutofit`` explicitly, the former removes every
        auto-fit choice so the setting is inherited from the layout /
        master / theme. Both paths must clear a prior ``a:spAutoFit``.
        """
        _, textbox = _make_textbox()
        tf = textbox.text_frame

        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        assert "<a:spAutoFit" in tf._txBody.xml

        tf.auto_size = None

        assert tf.auto_size is None
        xml = tf._txBody.xml
        assert "<a:spAutoFit" not in xml
        assert "<a:normAutofit" not in xml
        assert "<a:noAutofit" not in xml

    # -- 5: cross-reference to the adjacent issues ---------------------------

    def it_documents_the_cross_reference_to_issues_525_and_715(self):
        """Scenario 5: #970 sits adjacent to #525 and #715 on the tracker.

        * #525 is the textbox auto-grow ask #970 duplicates. Verified by
          ``tests/test_issue_525_shape_autofit_verify.py``, which pins
          every member of :class:`MSO_AUTO_SIZE` against its ``a:bodyPr``
          child (``a:spAutoFit`` / ``a:normAutofit`` / ``a:noAutofit``).
        * #715 is the placeholder text-shrink ask. Verified by the
          read/write :attr:`TextFrame.font_scale` and
          :attr:`TextFrame.line_space_reduction` properties, which
          expose the ``fontScale`` / ``lnSpcReduction`` attributes on
          ``a:normAutofit``.

        All three issues land on the same ``a:bodyPr`` auto-fit-choice
        machinery; this "documentation test" guarantees the public API
        surface #970 hinges on is still present and nominally hooked up
        to ``a:normAutofit``'s tuning attributes so the cross-reference
        to #715 remains valid.
        """
        # -- #970's answer: MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT on TextFrame --
        assert hasattr(TextFrame, "auto_size")
        assert MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT in MSO_AUTO_SIZE
        assert MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE in MSO_AUTO_SIZE
        assert MSO_AUTO_SIZE.NONE in MSO_AUTO_SIZE

        # -- #715's answer: font_scale / line_space_reduction on TextFrame --
        assert hasattr(TextFrame, "font_scale")
        assert hasattr(TextFrame, "line_space_reduction")

        # -- sanity: the three adjacent members co-exist on a single bodyPr --
        _, textbox = _make_textbox()
        tf = textbox.text_frame

        # -- #970/#525 path --
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        assert "<a:spAutoFit" in tf._txBody.xml

        # -- #715 path (requires a:normAutofit; writing font_scale must
        #    transparently swap spAutoFit for normAutofit, matching the
        #    EG_TextAutofit choice-group semantics) --
        tf.font_scale = 85.0
        xml = tf._txBody.xml
        assert "<a:normAutofit" in xml
        assert 'fontScale="85000"' in xml
        assert "<a:spAutoFit" not in xml
