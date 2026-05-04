# pyright: reportPrivateUsage=false

"""Regression test for issue #1099 — "Missing slide notes after slide 20".

Issue #1099 (https://github.com/scanny/python-pptx/issues/1099) reported
that extracting slide notes via
``slide.notes_slide.notes_text_frame.text`` returned an empty string for
slides past index 20 in a specific deck. The reporter attributed this to
"a known bug due to XML parsing" but never supplied the deck, a minimal
repro, or a citation for the "known bug" claim.

Inspecting the code path makes clear there is no index-dependent logic
involved:

* :attr:`.Slide.has_notes_slide` — iterates slide-part relationships
  looking for an ``RT.NOTES_SLIDE`` target. Independent of slide ordinal.
* :meth:`.Slide.notes_slide` / :meth:`.SlidePart.notes_slide` — resolves
  (or lazily creates) the notes slide part on the same relationship.
* :attr:`.NotesSlide.notes_placeholder` — iterates the notes-slide's
  placeholders looking for the one with ``ph_type == BODY``.
* :attr:`.NotesSlide.notes_text_frame` — shortcut to the notes
  placeholder's ``text_frame``.

None of these consult, or are constrained by, the slide's index in the
presentation's slide-id list. The "slide 20" threshold has no hook in
the library.

This suite pins the scenario end-to-end to close the issue as
not-reproducible:

  1. Author a 25-slide deck, put distinctive notes on every slide, save,
     reopen, and assert every slide — including every slide past index
     20 — round-trips its notes text exactly.
  2. Extend to 50 slides with a mix of short, long, and Unicode-heavy
     notes. Still no index-based drop-off.
  3. Verify :attr:`.Slide.has_notes_slide` is |True| for every slide with
     notes in the reopened deck, and that :attr:`.NotesSlide.notes_text_frame`
     is not |None| on any of them.

Any future regression that actually drops notes past an arbitrary slide
ordinal would be caught by every one of these.
"""

from __future__ import annotations

import io
from typing import TYPE_CHECKING

import pytest

from pptx import Presentation

if TYPE_CHECKING:
    from pptx.presentation import Presentation as PresentationCls


# -- helpers ---------------------------------------------------------------


def _note_for(i: int) -> str:
    """Distinctive deterministic notes string for slide index `i` (0-based)."""
    return f"Speaker notes for slide #{i + 1} — marker-{i + 1:04d}"


def _roundtrip(prs: PresentationCls) -> PresentationCls:
    """Save `prs` to an in-memory stream and reopen it."""
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


# -- scenarios -------------------------------------------------------------


class Describe_issue_1099:
    """Notes extraction works for every slide, including beyond slide 20."""

    def it_round_trips_notes_on_all_25_slides(self):
        prs = Presentation()
        blank = prs.slide_layouts[6]
        for i in range(25):
            slide = prs.slides.add_slide(blank)
            tf = slide.notes_slide.notes_text_frame
            assert tf is not None
            tf.text = _note_for(i)

        reopened = _roundtrip(prs)

        assert len(reopened.slides) == 25
        for i, slide in enumerate(reopened.slides):
            assert slide.has_notes_slide, f"slide #{i + 1} lost its notes relationship"
            tf = slide.notes_slide.notes_text_frame
            assert tf is not None, f"slide #{i + 1} has no notes text frame"
            assert tf.text == _note_for(i), f"slide #{i + 1} notes mismatched: {tf.text!r}"

    def it_round_trips_notes_past_slide_20_specifically(self):
        # -- explicit regression pin on the #1099-reported boundary
        prs = Presentation()
        blank = prs.slide_layouts[6]
        for i in range(25):
            slide = prs.slides.add_slide(blank)
            tf = slide.notes_slide.notes_text_frame
            assert tf is not None
            tf.text = _note_for(i)

        reopened = _roundtrip(prs)

        # -- every slide past index 20 (0-based) must round-trip cleanly
        for i in range(20, 25):
            slide = reopened.slides[i]
            assert slide.has_notes_slide
            tf = slide.notes_slide.notes_text_frame
            assert tf is not None
            assert tf.text == _note_for(i)

    def it_round_trips_varied_notes_on_a_50_slide_deck(self):
        """Longer deck with mixed short, long, and Unicode notes.

        Exercises the reporter's "specific deck" framing — if anything in
        the OPC / relationship / notes-slide-part machinery were
        index-dependent, a 50-slide mix would tease it out.
        """
        prs = Presentation()
        blank = prs.slide_layouts[6]

        def _mixed_note(i: int) -> str:
            if i % 5 == 0:
                # -- long note, many paragraphs worth of text
                return f"Long note #{i + 1} :: " + ("paragraph body " * 80)
            if i % 7 == 0:
                # -- Unicode-heavy
                return f"Note #{i + 1} — üñíçôdé 中文 テスト 🎯"
            return _note_for(i)

        for i in range(50):
            slide = prs.slides.add_slide(blank)
            tf = slide.notes_slide.notes_text_frame
            assert tf is not None
            tf.text = _mixed_note(i)

        reopened = _roundtrip(prs)

        assert len(reopened.slides) == 50
        for i, slide in enumerate(reopened.slides):
            assert slide.has_notes_slide, f"slide #{i + 1} lost its notes relationship"
            tf = slide.notes_slide.notes_text_frame
            assert tf is not None, f"slide #{i + 1} has no notes text frame"
            assert tf.text == _mixed_note(i), f"slide #{i + 1} notes diverged on round-trip"

    @pytest.mark.parametrize("n_slides", [21, 25, 30, 40])
    def it_has_notes_slide_is_true_past_slide_20(self, n_slides: int):
        prs = Presentation()
        blank = prs.slide_layouts[6]
        for i in range(n_slides):
            slide = prs.slides.add_slide(blank)
            tf = slide.notes_slide.notes_text_frame
            assert tf is not None
            tf.text = _note_for(i)

        reopened = _roundtrip(prs)

        assert all(s.has_notes_slide for s in reopened.slides)
        assert all(s.notes_slide.notes_text_frame is not None for s in reopened.slides)
