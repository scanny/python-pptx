# pyright: reportPrivateUsage=false

"""Regression test for issue #655 — "numbered list helper on a TextFrame".

Issue #655 (https://github.com/scanny/python-pptx/issues/655) asked for a
convenience on ``TextFrame`` for turning every paragraph in a text frame
into a numbered list entry — "I want my whole text-frame numbered, not to
call ``auto_number`` on each paragraph individually".

The existing per-paragraph API (``paragraph.bullet.auto_number(scheme,
start_at=None)`` — shipped under #100) already composes cleanly with a
``for p in text_frame.paragraphs: …`` loop, which is readable, flexible
(skip headings, interleave non-numbered paragraphs, etc.), and does not
require an opinionated new collection-level method whose semantics would
inevitably wobble (skip empty paragraphs? restart numbering? apply to
every paragraph or only un-numbered ones?).

#655 is therefore verified-and-closed via the documentation recipe added
in ``docs/user/text.rst`` — the "Numbered lists" subsection — plus the
existing ``_BulletFormat.auto_number`` API. This suite pins that the
recipe actually works end-to-end:

  * calling ``auto_number`` on every paragraph of a text frame writes an
    ``a:buAutoNum`` child under each paragraph's ``a:pPr`` with the
    requested scheme;
  * a ``start_at`` argument is threaded onto the first paragraph (and,
    passed uniformly, onto every paragraph — PowerPoint treats the first
    numbered paragraph's ``startAt`` as the anchor for the run);
  * the numbered state survives ``Presentation.save`` + reopen — the
    ``buAutoNum`` elements are present on the reloaded deck with their
    ``type`` and ``startAt`` attributes intact;
  * interleaving ``bullet.auto_number`` with ``bullet.none`` works — a
    spacer paragraph in the middle is explicitly un-numbered without
    disturbing the numbered paragraphs on either side.

Any future regression that drops the per-paragraph ``auto_number`` API,
or that corrupts ``a:buAutoNum`` serialization on save+reopen, will
reproduce the #655 reporter's pain here.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.enum.text import PP_AUTO_NUMBER_SCHEME
from pptx.oxml.ns import qn
from pptx.util import Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against test-module pollution of ``PartFactory.part_type_for``.

    Earlier test modules in the suite overwrite the part-type dispatch
    table with mocks and do not restore it; ``Presentation(stream)``
    reopen routes through that table, so a polluted dispatch makes the
    reopen branch fail spuriously. Mirrors the pattern used by
    ``tests/test_issue_1013_placeholder_color_verify.py``.
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


def _populated_text_frame(prs, *texts):
    """Return a text-frame on a freshly-added blank slide containing `texts`
    across one paragraph per entry.

    The first string lands on the pre-existing empty paragraph via
    ``text_frame.text =``; subsequent strings go in via ``add_paragraph()``
    so the frame ends up with exactly ``len(texts)`` paragraphs.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(3))
    tf = tb.text_frame
    tf.text = texts[0]
    for t in texts[1:]:
        tf.add_paragraph().text = t
    return tf


class DescribeIssue655NumberedListsVerify:
    """#655 numbered-list recipe verify-and-close.

    Pins that the ``for p in text_frame.paragraphs: p.bullet.auto_number(...)``
    idiom documented in ``docs/user/text.rst`` actually does what the #655
    reporter wanted: numbers every paragraph in a text frame.
    """

    # -- happy path: the documented recipe numbers every paragraph --------

    def it_numbers_every_paragraph_when_the_recipe_is_applied(self):
        """The canonical #655 scenario — three paragraphs, one scheme.

        Iterates over ``text_frame.paragraphs`` and calls
        ``bullet.auto_number(ARABIC_PERIOD)`` on each. Every paragraph must
        end up with an ``a:buAutoNum`` element whose ``type`` is the
        requested scheme.
        """
        prs = Presentation()
        tf = _populated_text_frame(prs, "First item", "Second item", "Third item")

        for paragraph in tf.paragraphs:
            paragraph.bullet.auto_number(PP_AUTO_NUMBER_SCHEME.ARABIC_PERIOD)

        assert len(tf.paragraphs) == 3
        for paragraph in tf.paragraphs:
            assert paragraph.bullet.type == "autonum"
            assert paragraph.bullet.number_scheme == PP_AUTO_NUMBER_SCHEME.ARABIC_PERIOD
            # -- no explicit start_at -> buAutoNum/@startAt absent; reader
            # -- falls back to PowerPoint's documented default of 1 --
            assert paragraph.bullet.start_at == 1

    def it_accepts_other_number_schemes(self):
        """Scheme-selection is not hard-wired to ARABIC_PERIOD.

        Re-runs the recipe with ``ROMAN_UC_PERIOD`` to pin that any
        ``PP_AUTO_NUMBER_SCHEME`` member flows through the loop idiom.
        """
        prs = Presentation()
        tf = _populated_text_frame(prs, "I", "II", "III")

        for paragraph in tf.paragraphs:
            paragraph.bullet.auto_number(PP_AUTO_NUMBER_SCHEME.ROMAN_UC_PERIOD)

        for paragraph in tf.paragraphs:
            assert paragraph.bullet.number_scheme == PP_AUTO_NUMBER_SCHEME.ROMAN_UC_PERIOD

    def it_threads_start_at_onto_every_paragraph(self):
        """``start_at`` is propagated when passed uniformly.

        The docs recipe shows passing the same ``start_at`` to every
        paragraph; doing so must not collide or zero out. Each paragraph
        should record the requested start ordinal.
        """
        prs = Presentation()
        tf = _populated_text_frame(prs, "Fifth", "Sixth", "Seventh")

        for paragraph in tf.paragraphs:
            paragraph.bullet.auto_number(
                PP_AUTO_NUMBER_SCHEME.ARABIC_PERIOD, start_at=5
            )

        for paragraph in tf.paragraphs:
            assert paragraph.bullet.number_scheme == PP_AUTO_NUMBER_SCHEME.ARABIC_PERIOD
            assert paragraph.bullet.start_at == 5

    # -- skipping a heading paragraph --------------------------------------

    def it_lets_the_caller_skip_a_heading_paragraph(self):
        """The docs recipe's "skip paragraph 0" variant.

        Loop with an index, branch on ``i == 0``. The heading paragraph
        carries no bullet; the remaining paragraphs are numbered.
        """
        prs = Presentation()
        tf = _populated_text_frame(prs, "Heading", "Item A", "Item B")

        for i, paragraph in enumerate(tf.paragraphs):
            if i == 0:
                continue
            paragraph.bullet.auto_number(PP_AUTO_NUMBER_SCHEME.ARABIC_PERIOD)

        assert tf.paragraphs[0].bullet.type is None
        assert tf.paragraphs[1].bullet.type == "autonum"
        assert tf.paragraphs[2].bullet.type == "autonum"

    # -- interleaving numbered and un-numbered paragraphs ------------------

    def it_supports_interleaving_numbered_and_unnumbered_paragraphs(self):
        """The ``bullet.none()`` spacer variant from the recipe.

        Put an explicit-no-bullet paragraph between two numbered ones.
        The spacer carries a ``a:buNone`` element; the numbered siblings
        carry ``a:buAutoNum`` as usual.
        """
        prs = Presentation()
        tf = _populated_text_frame(prs, "One", "", "Two")

        tf.paragraphs[0].bullet.auto_number(PP_AUTO_NUMBER_SCHEME.ARABIC_PERIOD)
        tf.paragraphs[1].bullet.none()
        tf.paragraphs[2].bullet.auto_number(PP_AUTO_NUMBER_SCHEME.ARABIC_PERIOD)

        assert tf.paragraphs[0].bullet.type == "autonum"
        assert tf.paragraphs[1].bullet.type == "none"
        assert tf.paragraphs[2].bullet.type == "autonum"

    # -- round-trip through save + reopen ----------------------------------

    def it_round_trips_the_numbered_list_through_save_and_reload(
        self, _restore_part_factory
    ):
        """The #655 reporter will save the deck; numbering must survive.

        Applies the recipe, saves to a ``.pptx`` byte stream, reopens, and
        asserts each reopened paragraph still carries an ``a:buAutoNum``
        with the right scheme.
        """
        prs = Presentation()
        tf = _populated_text_frame(prs, "First", "Second", "Third")

        for paragraph in tf.paragraphs:
            paragraph.bullet.auto_number(
                PP_AUTO_NUMBER_SCHEME.ARABIC_PERIOD, start_at=2
            )

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        # -- locate the textbox by its text content (layout 5 also carries
        # -- a title placeholder, so we can't rely on shape index) --
        tf2 = None
        for shape in prs2.slides[0].shapes:
            if shape.has_text_frame and shape.text_frame.paragraphs[0].text == "First":
                tf2 = shape.text_frame
                break
        assert tf2 is not None

        assert [p.text for p in tf2.paragraphs] == ["First", "Second", "Third"]
        for paragraph in tf2.paragraphs:
            assert paragraph.bullet.number_scheme == PP_AUTO_NUMBER_SCHEME.ARABIC_PERIOD
            assert paragraph.bullet.start_at == 2

    # -- XML-shape invariant: each pPr has exactly one buAutoNum -----------

    def it_writes_exactly_one_buAutoNum_per_paragraph(self):
        """Pin the XML: no leftover bullet choice siblings.

        ``set_auto_number_bullet`` must replace any existing bullet-choice
        child (``a:buNone``, ``a:buChar``, or a prior ``a:buAutoNum``). A
        regression that appended rather than replaced would leave dead
        siblings in the ``a:pPr`` that PowerPoint might or might not
        honour consistently.
        """
        prs = Presentation()
        tf = _populated_text_frame(prs, "A", "B")

        # -- pre-seed a different bullet so we can see the replacement --
        tf.paragraphs[0].bullet.character("•")
        for paragraph in tf.paragraphs:
            paragraph.bullet.auto_number(PP_AUTO_NUMBER_SCHEME.ARABIC_PERIOD)

        for paragraph in tf.paragraphs:
            pPr = paragraph._pPr
            assert pPr is not None
            assert pPr.find(qn("a:buChar")) is None
            assert pPr.find(qn("a:buNone")) is None
            buAutoNum_children = pPr.findall(qn("a:buAutoNum"))
            assert len(buAutoNum_children) == 1

    # -- API surface guard -------------------------------------------------

    def it_exposes_the_auto_number_API_the_recipe_relies_on(self):
        """Guard: the per-paragraph ``bullet.auto_number`` API is present.

        If a refactor ever removed or renamed ``_BulletFormat.auto_number``
        the documented #655 recipe would silently break. This test catches
        that — the API surface the user-guide references must exist.
        """
        from pptx.text.text import _BulletFormat

        assert hasattr(_BulletFormat, "auto_number")
        assert hasattr(_BulletFormat, "none")
        assert hasattr(_BulletFormat, "clear")
