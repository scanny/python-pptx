# pyright: reportPrivateUsage=false

"""Regression test for issue #1026 — `fit_text()` crashes on special chars.

Reporter (https://github.com/scanny/python-pptx/issues/1026) saw
``TextFrame.fit_text()`` raise an opaque crash when the shape text included
glyphs the selected font could not measure — e.g. arrows (``↑``, ``↓``,
``→``, ``←``), CJK ideographs, or emoji. Pillow's ``ImageFont.getbbox`` /
``ImageFont.getlength`` can raise ``UnicodeEncodeError`` / ``OSError`` /
``ValueError`` for such glyphs depending on the Pillow/FreeType version
installed, which previously surfaced as an uncaught exception from deep
inside the layout loop.

The fix (``src/pptx/text/layout.py::_rendered_size``) falls back to
per-character measurement and substitutes the width of ``?`` for any
character the font cannot measure, so `fit_text()` returns a sensible
integer point size rather than crashing. When the font itself fails to
open (already covered by #168), the caller still sees an actionable
:class:`pptx.exc.TextLayoutError`.
"""

from __future__ import annotations

from os.path import abspath, dirname, join
from typing import cast

import pytest

from pptx import Presentation
from pptx.text import layout as layout_module
from pptx.text.layout import TextFitter, _rendered_size
from pptx.util import Inches

FONT_FILE = abspath(
    join(
        dirname(__file__),
        "..",
        "features",
        "steps",
        "test_files",
        "calibriz.ttf",
    )
)


def _text_frame_with(text: str):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    tb.text_frame.text = text
    return tb.text_frame


class DescribeIssue1026FitTextSpecialChars(object):
    """Regression scenarios for #1026 fit_text crashes on special chars."""

    # ------------------------------------------------------------------
    # end-to-end: fit_text() must succeed, not crash, on special chars
    # ------------------------------------------------------------------

    @pytest.mark.parametrize(
        "text",
        [
            "→ Hello ← Arrow",  # arrows surrounded by spaces
            "580.9m↓-11.3m592.2m",  # reporter's exact text (no spaces around arrow)
            "Hello 你好 world",  # CJK with spaces
            "\U0001f389 Party \U0001f600",  # emoji
            "Mixed ← → ↑ ↓ 你 好",  # mixed
        ],
    )
    def it_fits_text_containing_special_chars_without_crashing(self, text: str):
        # -- with Calibri (which supports arrows and many Unicode chars)
        # -- fit_text() should return cleanly.
        text_frame = _text_frame_with(text)
        text_frame.fit_text(max_size=48, font_file=FONT_FILE)

        assert text_frame.word_wrap is True
        sizes = [
            cast(float, run.font.size.pt)
            for paragraph in text_frame.paragraphs
            for run in paragraph.runs
            if run.font.size is not None
        ]
        assert sizes, "fit_text should have applied a size to at least one run"
        assert all(1 <= size <= 48 for size in sizes)

    # ------------------------------------------------------------------
    # unit: `_rendered_size` falls back gracefully when Pillow raises
    # ------------------------------------------------------------------

    def it_falls_back_to_per_char_measurement_on_UnicodeEncodeError(
        self, monkeypatch: pytest.MonkeyPatch
    ):
        """When `_measure_px` raises `UnicodeEncodeError` for the full text,
        the fallback measures character-by-character, substituting `?` for
        any char that still fails. Covers the Pillow "font cannot encode this
        glyph" path reported in #1026.
        """
        calls: list[str] = []
        real_measure = layout_module._measure_px

        def fake_measure(font: object, text: str):
            calls.append(text)
            # -- simulate Pillow raising for the multi-char call containing
            # -- an arrow, but succeeding for single-char measurements.
            if len(text) > 1 and "→" in text:
                raise UnicodeEncodeError("utf-8", text, 0, 1, "missing glyph")
            if text == "→":
                raise UnicodeEncodeError("utf-8", text, 0, 1, "missing glyph")
            return real_measure(font, text)

        monkeypatch.setattr(layout_module, "_measure_px", fake_measure)

        w, h = _rendered_size("A→B", 12, FONT_FILE)

        # -- must return sensible positive values, not propagate the error --
        assert w > 0
        assert h > 0
        # -- the full-string call raised, then single-char path ran: at
        # -- minimum the full string and a fallback measurement are present.
        assert "A→B" in calls
        assert "?" in calls  # fallback glyph was used at least once

    def it_falls_back_to_per_char_measurement_on_OSError(self, monkeypatch: pytest.MonkeyPatch):
        """Same fallback kicks in when Pillow raises `OSError` (seen on some
        Pillow/FreeType builds for complex text layout failures).
        """
        real_measure = layout_module._measure_px

        def fake_measure(font: object, text: str):
            if len(text) > 1:
                raise OSError("raqm: shape failure")
            return real_measure(font, text)

        monkeypatch.setattr(layout_module, "_measure_px", fake_measure)

        w, h = _rendered_size("Hello", 12, FONT_FILE)

        assert w > 0
        assert h > 0

    def it_falls_back_to_per_char_measurement_on_ValueError(self, monkeypatch: pytest.MonkeyPatch):
        """Same fallback for `ValueError` (Pillow raises this for some
        surrogate/unassigned-codepoint cases).
        """
        real_measure = layout_module._measure_px

        def fake_measure(font: object, text: str):
            if len(text) > 1:
                raise ValueError("invalid character sequence")
            return real_measure(font, text)

        monkeypatch.setattr(layout_module, "_measure_px", fake_measure)

        w, h = _rendered_size("Hello", 12, FONT_FILE)

        assert w > 0
        assert h > 0

    def it_substitutes_question_mark_width_for_unmeasurable_chars(
        self, monkeypatch: pytest.MonkeyPatch
    ):
        """A char that can't be measured is treated as `?` (the OpenType-
        required `.notdef`-fallback glyph), producing a non-zero width.
        """
        real_measure = layout_module._measure_px

        def fake_measure(font: object, text: str):
            # -- bytes-like hack: only the unmeasurable codepoint fails,
            # -- forcing the fallback to use `?` in its place.
            if text == "�":
                raise ValueError("unmeasurable")
            return real_measure(font, text)

        monkeypatch.setattr(layout_module, "_measure_px", fake_measure)

        # -- width of "A�" with fallback == width of "A" + width of "?" --
        fallback_w, _ = _rendered_size("A�", 12, FONT_FILE)
        reference_w, _ = _rendered_size("A?", 12, FONT_FILE)

        assert fallback_w == reference_w

    # ------------------------------------------------------------------
    # `TextFitter.best_fit_font_size` still produces integer result
    # ------------------------------------------------------------------

    def it_returns_integer_point_size_for_text_with_special_chars(self):
        size = TextFitter.best_fit_font_size("→ Hello ← Arrow", (4000000, 1000000), 48, FONT_FILE)
        assert isinstance(size, int)
        assert 1 <= size <= 48

    def it_returns_integer_point_size_for_reporter_exact_text(self):
        # -- exact text and small shape from the #1026 report --
        size = TextFitter.best_fit_font_size(
            "580.9m↓-11.3m592.2m", (4000000, 1000000), 48, FONT_FILE
        )
        assert isinstance(size, int)
        assert 1 <= size <= 48
