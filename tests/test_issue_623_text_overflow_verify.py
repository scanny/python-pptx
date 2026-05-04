# pyright: reportPrivateUsage=false

"""Verify-close regression test for issue #623 — text overflow control.

Issue #623 (https://github.com/scanny/python-pptx/issues/623) asked for a
way to control how text overflows the bounds of a PowerPoint textbox.
PowerPoint exposes three behaviours:

* (a) **Wrap** — text wraps at the shape's right-hand edge
  (``a:bodyPr/@wrap="square"``).
* (b) **No wrap** — text lays out on a single line and overflows the
  shape horizontally (``a:bodyPr/@wrap="none"``).
* (c) **Clip** — with ``<a:noAutofit/>`` (``MSO_AUTO_SIZE.NONE``), any
  text that still exceeds the shape's extent after layout is clipped by
  the shape boundary; PowerPoint does not resize text or shape to
  accommodate the overflow.

No wrapper API is required. Every primitive already exists on the
public surface:

* :attr:`TextFrame.word_wrap` — ``True`` / ``False`` / ``None`` round-trip
  onto ``a:bodyPr/@wrap="square"`` / ``"none"`` / attribute-absent.
* :attr:`TextFrame.auto_size` — ``MSO_AUTO_SIZE.NONE`` emits
  ``<a:noAutofit/>`` under ``a:bodyPr``, which is PowerPoint's "clip
  overflow" mode.

This suite pins the three documented recipes so a regression that
breaks either the ``@wrap`` round-trip or the ``<a:noAutofit/>`` emit
fails fast and points at #623.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches


class DescribeIssue623TextOverflowRecipe:
    """Verify-close regression suite for the text-overflow recipes (issue #623)."""

    def it_wraps_text_at_the_shape_boundary_when_word_wrap_is_true(self):
        # -- Recipe (a): TextFrame.word_wrap = True emits
        # -- a:bodyPr/@wrap="square" so PowerPoint wraps text at the
        # -- shape's right edge.
        _, tf = _fresh_textbox()

        tf.word_wrap = True

        assert tf.word_wrap is True
        assert tf._txBody.bodyPr.get("wrap") == "square"

    def it_disables_wrapping_so_text_overflows_on_a_single_line(self):
        # -- Recipe (b): TextFrame.word_wrap = False emits
        # -- a:bodyPr/@wrap="none" so PowerPoint lays text out on a
        # -- single line that overflows the shape horizontally.
        _, tf = _fresh_textbox()

        tf.word_wrap = False

        assert tf.word_wrap is False
        assert tf._txBody.bodyPr.get("wrap") == "none"

    def it_clips_overflowing_text_when_auto_size_is_none(self):
        # -- Recipe (c): TextFrame.auto_size = MSO_AUTO_SIZE.NONE emits
        # -- <a:noAutofit/> under a:bodyPr. PowerPoint renders text at
        # -- its authored size and simply clips anything outside the
        # -- shape's bounds — no shape resize, no font scale.
        _, tf = _fresh_textbox()

        tf.auto_size = MSO_AUTO_SIZE.NONE

        assert tf.auto_size == MSO_AUTO_SIZE.NONE
        assert tf._txBody.bodyPr.find(_q("a:noAutofit")) is not None

    def it_returns_to_inherited_wrap_when_word_wrap_is_none(self):
        # -- Assigning None removes the @wrap attribute, so the
        # -- effective value is inherited from the style hierarchy.
        _, tf = _fresh_textbox()
        tf.word_wrap = True
        assert tf._txBody.bodyPr.get("wrap") == "square"

        tf.word_wrap = None

        assert tf.word_wrap is None
        assert "wrap" not in tf._txBody.bodyPr.attrib

    def it_round_trips_all_three_recipes_through_save_and_reopen(self):
        # -- End-to-end pin: after save + reopen, both @wrap="none"
        # -- and <a:noAutofit/> survive. This is the pin against a
        # -- serializer change that silently drops either.
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # -- shape 0: wrap=square (the default after explicit set)
        tf_wrap = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.5), Inches(2), Inches(1)
        ).text_frame
        tf_wrap.text = "wrap"
        tf_wrap.word_wrap = True

        # -- shape 1: wrap=none + auto_size=NONE (the "clip" recipe)
        tf_clip = slide.shapes.add_textbox(
            Inches(0.5), Inches(2), Inches(2), Inches(1)
        ).text_frame
        tf_clip.text = "clip"
        tf_clip.word_wrap = False
        tf_clip.auto_size = MSO_AUTO_SIZE.NONE

        reloaded = _roundtrip(prs)
        reloaded_shapes = [
            s for s in reloaded.slides[0].shapes if s.has_text_frame
        ]

        # -- shape 0 keeps wrap=True through the round-trip.
        assert reloaded_shapes[0].text_frame.word_wrap is True
        # -- shape 1 keeps both the wrap=False and auto_size=NONE legs.
        assert reloaded_shapes[1].text_frame.word_wrap is False
        assert reloaded_shapes[1].text_frame.auto_size == MSO_AUTO_SIZE.NONE

    def it_rejects_non_boolean_word_wrap_values(self):
        # -- Defensive pin: word_wrap only accepts True / False / None.
        # -- A string value must not silently land in the XML and
        # -- confuse PowerPoint at open time.
        import pytest

        _, tf = _fresh_textbox()

        with pytest.raises(ValueError):
            tf.word_wrap = "square"  # type: ignore[assignment]


# -- helpers --------------------------------------------------------------


def _fresh_textbox():
    """Return ``(prs, text_frame)`` — a new textbox on a blank slide."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
    return prs, tb.text_frame


def _q(tag: str) -> str:
    """Return Clark-notation qname for a prefixed tag."""
    from pptx.oxml.ns import qn

    return qn(tag)


def _roundtrip(prs):
    """Serialize `prs` to a BytesIO and reopen it. Returns the reloaded Presentation."""
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)
