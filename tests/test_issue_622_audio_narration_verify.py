# pyright: reportPrivateUsage=false

"""Verify-close regression test for issue #622 — per-slide audio narration.

Issue #622 (https://github.com/scanny/python-pptx/issues/622) asked for a
way to attach audio narration to a slide: the clip should start
automatically, the speaker icon should be hidden from the audience, and
the slide should advance when the narration ends.

No wrapper API is required. Every primitive already exists on the public
surface:

* :meth:`SlideShapes.add_movie` with ``mime_type="audio/*"`` and
  ``autoplay=True`` embeds the audio and wires the ``withPrevious``
  start condition.
* :attr:`BaseShape.is_hidden` marks the speaker icon hidden (``cNvPr``
  ``hidden="1"``).
* :attr:`Transition.advance_after_time` carries the slide's auto-advance
  delay (``p:transition/@advTm``, milliseconds).

This suite pins the *recipe* documented in ``docs/user/media.rst`` so a
regression that breaks autoplay wiring, the hidden attribute on a movie
shape, or the ``advTm`` round-trip fails fast and points at #622.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.shapes.picture import Movie
from pptx.util import Emu


class DescribeIssue622AudioNarrationRecipe:
    """Verify-close regression suite for the audio-narration recipe (issue #622)."""

    def it_embeds_an_audio_clip_with_autoplay_start_condition(self):
        # -- the first leg of the #622 recipe: add_movie with an audio
        # -- mime_type and autoplay=True returns a Movie whose
        # -- start_condition is "withPrevious" (PowerPoint's
        # -- Start: Automatically).
        _, slide = _fresh_slide()

        narration = _add_narration(slide)

        assert isinstance(narration, Movie)
        assert narration.start_condition == "withPrevious"

    def it_hides_the_speaker_icon_from_the_audience(self):
        # -- the second leg: setting is_hidden on the returned shape
        # -- flips the cNvPr hidden attribute so PowerPoint omits the
        # -- icon during slideshow.
        _, slide = _fresh_slide()
        narration = _add_narration(slide)

        narration.is_hidden = True

        assert narration.is_hidden is True

    def it_drives_slide_auto_advance_via_transition_advance_after_time(self):
        # -- the third leg: the slide's Transition carries the
        # -- auto-advance delay (in ms). Callers are expected to derive
        # -- this from the clip duration themselves — the library
        # -- deliberately does not pull in a media-duration dep.
        _, slide = _fresh_slide()

        slide.transition.advance_after_time = 5000  # 5.0 seconds

        assert slide.transition.advance_after_time == 5000

    def it_round_trips_the_full_narration_recipe_through_save_and_reopen(self):
        # -- end-to-end pin: after Presentation.save + reopen, all three
        # -- legs are still in effect on the reloaded shape. This is the
        # -- pin against a serializer change that silently drops any of
        # -- autoplay wiring, cNvPr hidden, or the transition advTm
        # -- attribute.
        prs, slide = _fresh_slide()
        narration = _add_narration(slide)
        narration.is_hidden = True
        slide.transition.advance_after_time = 5000

        reloaded = _roundtrip(prs)

        reloaded_slide = reloaded.slides[0]
        # -- locate the narration shape by media_type, not shape index,
        # -- since a blank layout may include latent placeholders.
        movies = [s for s in reloaded_slide.shapes if isinstance(s, Movie)]
        assert len(movies) == 1
        reloaded_movie = movies[0]
        assert reloaded_movie.start_condition == "withPrevious"
        assert reloaded_movie.is_hidden is True
        assert reloaded_slide.transition.advance_after_time == 5000

    def it_accepts_a_filesystem_path_as_the_audio_file(self):
        # -- the recipe uses the same add_movie signature as a video
        # -- clip; a str path is accepted without further conversion.
        _, slide = _fresh_slide()

        icon = Emu(228600)  # 0.25 in
        narration = slide.shapes.add_movie(
            "features/steps/test_files/silence.wav",
            Emu(0),
            Emu(0),
            icon,
            icon,
            mime_type="audio/wav",
            autoplay=True,
        )

        assert isinstance(narration, Movie)
        # -- the audio clip is wired for playback, not placed as a
        # -- manual-click video.
        assert narration.start_condition == "withPrevious"


# -- helpers --------------------------------------------------------------


def _fresh_slide():
    """Return ``(prs, slide)`` — a new ``Presentation`` with one blank slide."""
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    return prs, slide


def _add_narration(slide) -> Movie:
    """Add the canonical narration clip per the #622 recipe and return it."""
    icon = Emu(228600)  # 0.25 in
    shape = slide.shapes.add_movie(
        "features/steps/test_files/silence.wav",
        Emu(0),
        Emu(0),
        icon,
        icon,
        mime_type="audio/wav",
        autoplay=True,
    )
    assert isinstance(shape, Movie)
    return shape


def _roundtrip(prs):
    """Serialize `prs` to a BytesIO and reopen it. Returns the reloaded Presentation."""
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)
