# pyright: reportPrivateUsage=false

"""Regression tests for FU-1 — `_FontColorFormat` cache staleness after promote.

Follow-up to issue #1111. The deferred-promotion `_FontColorFormat` proxy is
cached on `font.__dict__["color"]` by :class:`~pptx.util.lazyproperty`. Before
the fix, the first `font.color.rgb = RGBColor(...)` assignment created the
`a:solidFill` but left a stale proxy (whose inner `_color` remained the
pre-promotion `_NoneColor`) in that cache slot. Subsequent reads of
`font.color.rgb` raised ``AttributeError: no .rgb property on color type
'_NoneColor'``.

The fix: `_FontColorFormat._promote()` drops the cached `color` entry on
`font.__dict__` and returns the live `ColorFormat` (i.e.
`font.fill.fore_color`) so the setter can apply the change to the proxy that
future reads will observe.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR
from pptx.util import Inches


@pytest.fixture
def run_in_textbox():
    """Yield ``(presentation, run)`` -- a fresh run on a blank-layout slide.

    Uses :meth:`_Paragraph.add_run` so the run starts with no `<a:rPr>` and
    therefore no `<a:solidFill>` -- the precondition that triggers the
    deferred-promotion `_FontColorFormat` path. The presentation is returned
    alongside the run so save-and-reopen scenarios can serialize without
    walking the parent chain.
    """
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = "Hello"
    return prs, run


class DescribeFontColorCachePostPromote:
    """`font.color` reads correctly after a first `.rgb`/`.theme_color` write."""

    def it_reads_rgb_after_initial_rgb_write_through_held_font(self, run_in_textbox):
        # -- Holding the Font reference is the key: without the fix, the
        # -- cached `_FontColorFormat` on `font.__dict__["color"]` carries a
        # -- stale `_NoneColor` inner, and the second `.rgb` read raises.
        _, run = run_in_textbox
        font = run.font

        font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

        assert font.color.rgb == RGBColor(0xFF, 0x00, 0x00)

    def it_reads_rgb_after_initial_rgb_write_through_fresh_font_access(self, run_in_textbox):
        # -- `Run.font` is a plain `@property` so each access constructs a new
        # -- Font, masking the cache-staleness bug. Keep this scenario as a
        # -- sanity check that the chained-access path still works.
        _, run = run_in_textbox

        run.font.color.rgb = RGBColor(0x00, 0xFF, 0x00)

        assert run.font.color.rgb == RGBColor(0x00, 0xFF, 0x00)

    def it_reads_color_type_after_initial_rgb_write(self, run_in_textbox):
        _, run = run_in_textbox
        font = run.font

        font.color.rgb = RGBColor(0x12, 0x34, 0x56)

        assert font.color.type == MSO_COLOR_TYPE.RGB

    def it_reads_theme_color_after_initial_theme_color_write(self, run_in_textbox):
        _, run = run_in_textbox
        font = run.font

        font.color.theme_color = MSO_THEME_COLOR.ACCENT_2

        assert font.color.theme_color == MSO_THEME_COLOR.ACCENT_2
        assert font.color.type == MSO_COLOR_TYPE.SCHEME

    def it_reflects_a_second_rgb_write_on_later_reads(self, run_in_textbox):
        # -- double-write path: first write promotes to solidFill; second
        # -- write goes through the regular ColorFormat setter.
        _, run = run_in_textbox
        font = run.font

        font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        font.color.rgb = RGBColor(0x00, 0x00, 0xFF)

        assert font.color.rgb == RGBColor(0x00, 0x00, 0xFF)

    def it_round_trips_a_set_rgb_through_save_and_reopen(self, run_in_textbox):
        # -- full user-visible scenario: XML is persisted, re-parsed, and the
        # -- color is still readable.
        prs, run = run_in_textbox
        run.font.color.rgb = RGBColor(0xAB, 0xCD, 0xEF)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        run2 = prs2.slides[0].shapes[-1].text_frame.paragraphs[0].runs[0]
        assert run2.font.color.rgb == RGBColor(0xAB, 0xCD, 0xEF)

    def it_keeps_font_color_cache_consistent_after_promote(self, run_in_textbox):
        # -- Explicitly verify that the two references to `font.color`
        # -- taken before and after the first write both now reflect the
        # -- solid-fill state.
        _, run = run_in_textbox
        font = run.font
        before = font.color  # caches a `_FontColorFormat` in font.__dict__

        before.rgb = RGBColor(0x11, 0x22, 0x33)
        after = font.color

        # -- after the write, the cached proxy must report the new color
        assert after.rgb == RGBColor(0x11, 0x22, 0x33)
        assert after.type == MSO_COLOR_TYPE.RGB
