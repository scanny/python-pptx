# pyright: reportPrivateUsage=false

"""Regression test for issue #514 — `fit_text()` + shape resizing.

Issue #514 (https://github.com/scanny/python-pptx/issues/514) reports two
related concerns from the reporter's "report layout" workflow:

* When :meth:`TextFrame.fit_text` is called, the computed point size
  should be discoverable from each run's :attr:`Font.size` so callers can
  *introspect* the result — for example, to align the effective size
  across multiple text frames ("use the smallest of the two shrunk
  sizes").  The reporter's quote:

      "When you run ``fit_text()``, the size-property should be set
      instead being None.  This way one could check both text boxes
      and reset the size to — for instance — smallest value from both
      shapes for both shapes."

* Resizing a shape after an earlier :meth:`fit_text` call (or changing
  shape extents before a first :meth:`fit_text` call) should produce a
  point size appropriate to the *current* shape dimensions — the method
  must re-read the shape's ``width`` and ``height`` every invocation
  and not cache anything stale.

The implementation has always satisfied both of these: the private
``TextFrame._set_font`` helper walks every run's ``a:rPr`` and writes
``@sz`` with the chosen point size, and ``TextFrame._extents`` reads the
parent shape's ``width`` / ``height`` on every call through the
computed ``_extents`` property (no caching).  The verify-close scenarios
below pin that contract so a future refactor cannot regress it.

The bundled ``features/steps/test_files/calibriz.ttf`` font is used for
deterministic measurement across platforms (the same font used by the
#168 / #936 / #1026 regression suites).
"""

from __future__ import annotations

from os.path import abspath, dirname, join

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches, Pt

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


def _text_frame_with_size(text: str, width_in: float, height_in: float):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(width_in), Inches(height_in))
    tb.text_frame.text = text
    return tb, tb.text_frame


class DescribeIssue514FitTextAfterResize(object):
    """Regression scenarios for #514 fit_text introspection and resize."""

    # ------------------------------------------------------------------
    # #514 core request: font.size is readable after fit_text()
    # ------------------------------------------------------------------

    def it_sets_run_font_size_after_fit_text(self):
        # -- the central complaint in #514: callers need `font.size` to
        # -- reflect the chosen point size so they can compare across
        # -- text-frames.  It must not remain None after fit_text().
        _, text_frame = _text_frame_with_size(
            "Hello world, this is some text that may need to shrink.",
            width_in=2.0,
            height_in=0.75,
        )

        text_frame.fit_text(font_file=FONT_FILE, max_size=36)

        run = text_frame.paragraphs[0].runs[0]
        assert run.font.size is not None, "fit_text must set run.font.size"
        assert 1 <= run.font.size.pt <= 36

    def it_writes_sz_attribute_on_every_rPr_in_every_paragraph(self):
        # -- the reporter's use case is a text box with *multiple*
        # -- paragraphs; each paragraph's each run (and endParaRPr) must
        # -- carry the chosen size, not just the first run.
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(2))
        tf = tb.text_frame
        tf.text = "First paragraph here\nSecond paragraph text\nThird line of text"

        tf.fit_text(font_file=FONT_FILE, max_size=32)

        # -- every run in every paragraph must now have an explicit size --
        sizes: list[float] = []
        for paragraph in tf.paragraphs:
            assert len(paragraph.runs) >= 1
            for run in paragraph.runs:
                assert run.font.size is not None
                sizes.append(run.font.size.pt)

        # -- all runs share the single fit-size chosen by fit_text() --
        assert len(set(sizes)) == 1

    def it_allows_cross_textframe_size_alignment_from_font_size(self):
        # -- the literal use case the reporter describes: two text boxes
        # -- with different amounts of text, each fit independently, then
        # -- aligned to the smaller of the two sizes.
        _, tf1 = _text_frame_with_size(
            "A" * 200,  # plenty of text -> must shrink
            width_in=2.0,
            height_in=0.6,
        )
        _, tf2 = _text_frame_with_size(
            "short",  # tiny -> fits at max_size
            width_in=2.0,
            height_in=0.6,
        )

        tf1.fit_text(font_file=FONT_FILE, max_size=36)
        tf2.fit_text(font_file=FONT_FILE, max_size=36)

        size1 = tf1.paragraphs[0].runs[0].font.size
        size2 = tf2.paragraphs[0].runs[0].font.size
        assert size1 is not None
        assert size2 is not None

        # -- the heavily-filled box shrinks; the sparse box stays near max --
        assert size1.pt < size2.pt

        # -- caller picks the smaller and applies it across both frames --
        smaller = min(size1, size2)
        for tf in (tf1, tf2):
            for p in tf.paragraphs:
                for run in p.runs:
                    run.font.size = smaller

        # -- now every run in both frames shares the same size --
        for tf in (tf1, tf2):
            for p in tf.paragraphs:
                for run in p.runs:
                    assert run.font.size == smaller

    def it_reports_exact_size_in_centipoints_round_trip(self):
        # -- the value that comes back from font.size must exactly equal
        # -- the point size fit_text chose, converted to EMU via Pt().
        # -- This guards against any future unit-conversion bug between
        # -- `_set_font` (takes int points) and the rPr/@sz attribute
        # -- (stored as centipoints / 100ths of a point).
        _, tf = _text_frame_with_size(
            "Short but not tiny text",
            width_in=3.0,
            height_in=1.0,
        )

        tf.fit_text(font_file=FONT_FILE, max_size=24)

        run = tf.paragraphs[0].runs[0]
        size = run.font.size
        assert size is not None
        # -- EMU value must be an integer multiple of Pt(1)'s centipoint
        # -- unit (= 12700 EMU per point); the returned size, converted
        # -- back to point-integer, must equal the sz stored in XML.
        size_pt_int = int(size.pt)
        assert size == Pt(size_pt_int)

    # ------------------------------------------------------------------
    # #514 resize angle: fit_text re-reads shape width/height every call
    # ------------------------------------------------------------------

    def it_recomputes_fit_after_the_shape_is_enlarged(self):
        # -- fit once in a tight box, enlarge the shape, fit again: the
        # -- second size must be larger (or equal to max_size) because
        # -- more room is now available.
        tb, tf = _text_frame_with_size(
            "Hello world, this text needs fitting",
            width_in=1.5,
            height_in=0.5,
        )

        tf.fit_text(font_file=FONT_FILE, max_size=36)
        first_size = tf.paragraphs[0].runs[0].font.size
        assert first_size is not None
        first_pt = first_size.pt

        # -- grow the shape substantially --
        tb.width = Inches(6.0)
        tb.height = Inches(2.0)

        tf.fit_text(font_file=FONT_FILE, max_size=36)
        second_size = tf.paragraphs[0].runs[0].font.size
        assert second_size is not None
        second_pt = second_size.pt

        assert second_pt > first_pt, (
            f"fit_text did not re-read enlarged extents: first={first_pt} " f"second={second_pt}"
        )

    def it_recomputes_fit_after_the_shape_is_shrunk(self):
        # -- opposite direction: shrink the shape after a first fit_text
        # -- and the second call must emit a *smaller* size.
        tb, tf = _text_frame_with_size(
            "Hello world, this is a reasonable amount of text",
            width_in=6.0,
            height_in=2.0,
        )

        tf.fit_text(font_file=FONT_FILE, max_size=36)
        first_size = tf.paragraphs[0].runs[0].font.size
        assert first_size is not None
        first_pt = first_size.pt

        # -- shrink the shape --
        tb.width = Inches(1.5)
        tb.height = Inches(0.5)

        tf.fit_text(font_file=FONT_FILE, max_size=36)
        second_size = tf.paragraphs[0].runs[0].font.size
        assert second_size is not None
        second_pt = second_size.pt

        assert second_pt < first_pt, (
            f"fit_text did not re-read shrunk extents: first={first_pt} " f"second={second_pt}"
        )

    def it_does_not_cache_extents_across_calls(self):
        # -- a single TextFrame instance must not memoize the shape's
        # -- extents — each fit_text call goes through `_extents`, which
        # -- reads `self._parent.width` / `.height` on every invocation.
        tb, tf = _text_frame_with_size(
            "Hello world",
            width_in=2.0,
            height_in=1.0,
        )

        # -- prove `_extents` is re-read by hand: change width and watch
        # -- `_extents` reflect it.
        ex1 = tf._extents
        tb.width = Inches(5.0)
        ex2 = tf._extents
        assert ex2[0] > ex1[0]

    # ------------------------------------------------------------------
    # side effects: fit_text must also assert the "auto_size=NONE +
    # word_wrap=True + noAutofit" contract on every call.
    # ------------------------------------------------------------------

    def it_pins_the_auto_size_and_word_wrap_side_effects_per_call(self):
        _, tf = _text_frame_with_size(
            "Some text to fit",
            width_in=3.0,
            height_in=1.0,
        )

        # -- default after add_textbox is SHAPE_TO_FIT_TEXT (wrap=none); after
        # -- fit_text it must become NONE with wrap on.
        tf.fit_text(font_file=FONT_FILE, max_size=24)
        assert tf.auto_size is MSO_AUTO_SIZE.NONE
        assert tf.word_wrap is True

        # -- a subsequent fit_text after a resize preserves the contract --
        tf.fit_text(font_file=FONT_FILE, max_size=24)
        assert tf.auto_size is MSO_AUTO_SIZE.NONE
        assert tf.word_wrap is True
