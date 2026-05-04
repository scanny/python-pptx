"""Regression tests for issue #455 -- unified `Run.hyperlink` action surface.

The original reporter asked for `_Run.hyperlink` to support the full range
of actions that `ActionSetting` exposes on shapes -- slide jumps, screen
tips, sounds, and hyperlink-color control.

State before this fix: `_Run.hyperlink` carried `address` (URL) and (as of
issue #1077) `target_slide`, but lacked any way to set a ScreenTip or
attach a sound.  Color control was already available as
`Run.font.use_theme_hyperlink_color`.

State after this fix: `_Run.hyperlink` now also carries `screen_tip` and
`sound` / `set_sound` / `remove_sound`, so a run-level hyperlink has the
same capability surface as a shape's `click_action` for the run-valid
subset of actions (URL, slide jump, tooltip, embedded sound).  A run
`a:rPr` has no `a:hlinkMouseOver` child in python-pptx's schema model,
so there is no "hover" half for runs -- that stays shape-only.

These tests pin the end-to-end round-trip: set each piece of a run
hyperlink through the public API, save the presentation to a bytes
stream, reopen it, and verify the values come back.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.media import Audio
from pptx.util import Inches


def _add_textbox_with_run(prs, text="link"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = text
    return slide, tb, run


class DescribeIssue455RunHyperlinkUnifiedSurface(object):
    """Regression suite for issue #455 -- unified run hyperlink capabilities."""

    # -- address (URL) -- pre-existing, pinned for completeness ----------

    def it_round_trips_an_external_URL_on_a_run(self):
        prs = Presentation()
        _, _, run = _add_textbox_with_run(prs)
        run.hyperlink.address = "https://example.com/"

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        run2 = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        assert run2.hyperlink.address == "https://example.com/"

    # -- target_slide (slide jump, issue #1077) -- pinned here too -------

    def it_round_trips_a_slide_jump_on_a_run(self):
        prs = Presentation()
        slide1, _, run = _add_textbox_with_run(prs)
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        run.hyperlink.target_slide = slide2

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        run2 = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        assert run2.hyperlink.address is None
        assert run2.hyperlink.target_slide is not None
        assert run2.hyperlink.target_slide.slide_id == prs2.slides[1].slide_id

    # -- screen_tip (issue #455) -----------------------------------------

    def it_sets_a_screen_tip_creating_the_hlinkClick_element(self):
        prs = Presentation()
        _, _, run = _add_textbox_with_run(prs)

        run.hyperlink.screen_tip = "Hover me"

        rPr = run._r.rPr
        assert rPr is not None
        assert rPr.hlinkClick is not None
        assert rPr.hlinkClick.tooltip == "Hover me"

    def it_round_trips_a_URL_with_a_screen_tip(self):
        prs = Presentation()
        _, _, run = _add_textbox_with_run(prs)
        run.hyperlink.address = "https://example.com/"
        run.hyperlink.screen_tip = "Open example.com"

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        run2 = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        assert run2.hyperlink.address == "https://example.com/"
        assert run2.hyperlink.screen_tip == "Open example.com"

    def it_round_trips_a_slide_jump_with_a_screen_tip(self):
        prs = Presentation()
        slide1, _, run = _add_textbox_with_run(prs)
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        run.hyperlink.target_slide = slide2
        run.hyperlink.screen_tip = "Jump to slide 2"

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        run2 = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        assert run2.hyperlink.screen_tip == "Jump to slide 2"
        assert run2.hyperlink.target_slide is not None

    def it_clears_only_the_tooltip_when_screen_tip_is_set_to_None(self):
        prs = Presentation()
        _, _, run = _add_textbox_with_run(prs)
        run.hyperlink.address = "https://example.com/"
        run.hyperlink.screen_tip = "tip"

        run.hyperlink.screen_tip = None

        # -- tooltip gone, URL intact
        assert run.hyperlink.screen_tip is None
        assert run.hyperlink.address == "https://example.com/"

    # -- sound (issue #455) ----------------------------------------------

    def it_attaches_a_sound_to_a_run_hyperlink(self):
        prs = Presentation()
        _, _, run = _add_textbox_with_run(prs)
        audio = Audio.from_blob(b"RIFFWAVEfake", None, "applause.wav")

        sound = run.hyperlink.set_sound(audio)

        assert sound.name == "applause.wav"
        assert sound.rId
        assert run.hyperlink.sound is not None
        assert run.hyperlink.sound.name == "applause.wav"

    def it_round_trips_a_URL_with_a_sound(self):
        prs = Presentation()
        _, _, run = _add_textbox_with_run(prs)
        run.hyperlink.address = "https://example.com/"
        run.hyperlink.set_sound(Audio.from_blob(b"RIFFWAVEboing", None, "boing.wav"))

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        run2 = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        assert run2.hyperlink.address == "https://example.com/"
        assert run2.hyperlink.sound is not None
        assert run2.hyperlink.sound.name == "boing.wav"
        assert run2.hyperlink.sound.blob == b"RIFFWAVEboing"

    def it_replaces_an_existing_sound_when_set_sound_called_again(self):
        prs = Presentation()
        _, _, run = _add_textbox_with_run(prs)
        run.hyperlink.set_sound(Audio.from_blob(b"A", None, "old.wav"))

        run.hyperlink.set_sound(Audio.from_blob(b"B", None, "new.wav"))

        assert run.hyperlink.sound is not None
        assert run.hyperlink.sound.name == "new.wav"
        assert run.hyperlink.sound.blob == b"B"

    def it_removes_the_sound_while_preserving_the_URL(self):
        prs = Presentation()
        _, _, run = _add_textbox_with_run(prs)
        run.hyperlink.address = "https://example.com/"
        run.hyperlink.set_sound(Audio.from_blob(b"X", None, "ping.wav"))

        run.hyperlink.remove_sound()

        assert run.hyperlink.sound is None
        assert run.hyperlink.address == "https://example.com/"

    def it_is_a_noop_to_remove_sound_when_no_sound_is_set(self):
        prs = Presentation()
        _, _, run = _add_textbox_with_run(prs)

        run.hyperlink.remove_sound()  # -- should not raise

        assert run.hyperlink.sound is None

    def it_clears_audio_rel_when_clearing_the_hyperlink(self):
        """Clearing the URL also drops any embedded sound rel on the run.

        This mirrors `ActionSetting._clear_click_action`, which walks any
        `a:snd` child before removing the `a:hlinkClick` element so the
        AUDIO relationship it referenced is dropped.
        """
        prs = Presentation()
        slide, _, run = _add_textbox_with_run(prs)
        run.hyperlink.address = "https://example.com/"
        run.hyperlink.set_sound(Audio.from_blob(b"X", None, "ping.wav"))

        # -- count audio rels on the slide part before clearing
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT

        slide_part = slide.part
        audio_rels_before = [rel for rel in slide_part.rels.values() if rel.reltype == RT.AUDIO]
        assert audio_rels_before  # -- at least one

        run.hyperlink.address = None

        audio_rels_after = [rel for rel in slide_part.rels.values() if rel.reltype == RT.AUDIO]
        assert not audio_rels_after  # -- rel dropped too

    # -- font.use_theme_hyperlink_color interplay (issue #940) -----------

    def it_plays_nicely_with_use_theme_hyperlink_color(self):
        """The run's font colors are controlled separately via the
        `use_theme_hyperlink_color` opt-out.  Together with the now-unified
        hyperlink surface this gives runs the full `ActionSetting` capability
        set (URL / slide-jump / screen-tip / sound / color-override)."""
        from pptx.dml.color import RGBColor

        prs = Presentation()
        _, _, run = _add_textbox_with_run(prs)
        run.hyperlink.address = "https://example.com/"
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        run.font.use_theme_hyperlink_color = False

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        run2 = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        assert run2.hyperlink.address == "https://example.com/"
        assert run2.font.use_theme_hyperlink_color is False
