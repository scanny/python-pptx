# pyright: reportPrivateUsage=false

"""Regression test for issue #947 — ``shape.text_frame.text`` returns an
empty string for a shape whose paragraph contains an inline OMML math
equation wrapped in ``mc:AlternateContent``.

Issue #947 (https://github.com/scanny/python-pptx/issues/947) reports
that when a user inserts a math equation in PowerPoint, the resulting
``.pptx`` stores the equation as an ``mc:AlternateContent`` child of the
enclosing ``a:p`` — a ``mc:Choice`` subtree carrying the live
``a14:m/m:oMath`` (or ``m:oMathPara``) and a ``mc:Fallback`` subtree
carrying the plain-text rendering PowerPoint shows on pre-2010
consumers. ``_Paragraph.text`` previously iterated only the direct
``a:r`` / ``a:br`` / ``a:fld`` children of the paragraph, so the
fallback text was dropped and ``.text`` returned ``""``.

The fix (in ``src/pptx/text/text.py``) teaches the paragraph text
walker to drop into the ``mc:Fallback`` subtree of any
``mc:AlternateContent`` child and include its ``a:r`` / ``a:br`` /
``a:fld`` contents in document order. The live ``mc:Choice`` side is
intentionally ignored — its OMML rendering is what ``mc:Fallback``
already mirrors in plain text.

This test pins the behaviour from the caller's perspective: an
equation-bearing paragraph authored via ``_Paragraph.add_math_equation``
(which writes the same ``mc:AlternateContent`` scaffolding PowerPoint
does) must now surface the fallback text through
``TextFrame.text``, ``Paragraph.text``, and the shape-level ``.text``
accessor.
"""

from __future__ import annotations

from pptx import Presentation

_OMML = (
    "<m:oMathPara "
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    "<m:oMath>"
    "<m:r><m:t>E=mc</m:t></m:r>"
    "<m:sSup><m:e><m:r><m:t></m:t></m:r></m:e>"
    "<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>"
    "</m:oMath>"
    "</m:oMathPara>"
)


def _presentation_with_equation():
    """Return a freshly authored ``(Presentation, TextBox)`` whose single
    paragraph contains ``"Formula: "``, an inline math equation (wrapped
    in ``mc:AlternateContent`` just as PowerPoint would write it), and
    trailing ``" done."``.
    """
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    tb = slide.shapes.add_textbox(914400, 914400, 4000000, 1000000)
    p = tb.text_frame.paragraphs[0]
    p.add_run().text = "Formula: "
    p.add_math_equation(_OMML)
    p.add_run().text = " done."
    return prs, tb


class DescribeIssue947MathEquationTextExtraction(object):
    """The #947 flow: extract ``.text`` from a paragraph whose math
    equation was authored as ``mc:AlternateContent`` with a
    ``mc:Fallback`` run.
    """

    def it_includes_mc_fallback_text_in_text_frame_text(self):
        _, tb = _presentation_with_equation()

        # -- the fallback text ``"E=mc2"`` is the concatenation of every
        # -- ``m:t`` descendant of the OMML fragment; it must appear
        # -- verbatim between the bracketing plain-text runs.
        assert tb.text_frame.text == "Formula: E=mc2 done."

    def it_includes_mc_fallback_text_in_paragraph_text(self):
        _, tb = _presentation_with_equation()
        paragraph = tb.text_frame.paragraphs[0]

        assert paragraph.text == "Formula: E=mc2 done."

    def it_round_trips_the_fallback_text_through_save_and_reload(self, tmp_path):
        prs, _ = _presentation_with_equation()
        pptx_path = tmp_path / "issue-947.pptx"
        prs.save(str(pptx_path))

        prs2 = Presentation(str(pptx_path))
        shape = prs2.slides[0].shapes[-1]

        assert shape.has_text_frame is True
        assert shape.text_frame.text == "Formula: E=mc2 done."

    def and_it_works_when_the_equation_is_the_only_paragraph_content(
        self,
    ):
        """A paragraph that contains only an ``mc:AlternateContent``
        equation (no sibling runs) must still surface the fallback
        text — previously ``.text`` returned ``""`` in that case.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tb = slide.shapes.add_textbox(914400, 914400, 4000000, 1000000)
        p = tb.text_frame.paragraphs[0]
        p.add_math_equation(_OMML)

        assert tb.text_frame.text == "E=mc2"

    def but_it_does_not_double_count_when_both_choice_and_fallback_hold_text(
        self,
    ):
        """Sanity check: the live ``mc:Choice`` side (``a14:m/m:oMath``)
        is intentionally skipped; only ``mc:Fallback`` contributes. The
        OMML ``m:t`` children are *inside* ``mc:Choice`` but are **not**
        ``a:r``/``a:br``/``a:fld``, so they never reach the walker.
        """
        _, tb = _presentation_with_equation()

        # -- ``"E=mc"`` and ``"2"`` live in the OMML under mc:Choice;
        # -- ``"E=mc2"`` is the concatenated fallback. If we walked the
        # -- Choice subtree we'd see the OMML characters twice.
        text = tb.text_frame.text
        assert text.count("E=mc2") == 1
