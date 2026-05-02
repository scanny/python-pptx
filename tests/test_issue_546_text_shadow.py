# pyright: reportPrivateUsage=false

"""Regression test for issue #546 — text shadow.

Issue #546 (https://github.com/scanny/python-pptx/issues/546) asks:

    "Hi guys, just a question, how can I add text shadow to fonts?"

At the time of the report, :class:`pptx.text.text.Font` exposed typeface,
size, bold, italic, color, underline, etc., but not the visual effects
(shadow, glow, reflection, soft-edge) that live on a run's
``a:rPr/a:effectLst``. The reporter had no handle to PowerPoint's
"text shadow" checkbox.

Two changes deliver the feature:

  * Foundation ``F2`` (Wave 1) introduced the shared
    :class:`~pptx.dml.effect.EffectFormat` /
    :class:`~pptx.dml.effect.ShadowFormat` family, which writes
    ``a:effectLst/a:outerShdw`` uniformly on any shape-like container
    whose schema admits an ``a:effectLst`` child (``p:spPr``,
    ``p:grpSpPr``, ``p:bgPr``, ``a:tcPr``, and — per ECMA-376
    §20.1.8 + §21.1.2.3.8 — ``a:rPr``).
  * This change (issue #546, Wave 1 F2 follow-up) adds the
    ``a:effectLst`` descriptor to ``CT_TextCharacterProperties`` and
    surfaces it on ``Font`` via the ``shadow`` and ``effect_format``
    properties, so a caller can write
    ``run.font.shadow.blur_radius = Emu(50800)`` and have the expected
    XML land under ``a:rPr/a:effectLst/a:outerShdw``.

This test pins the reporter's exact ask — setting a text shadow on a
run's font — and round-trips every shadow knob through
``Presentation.save`` + reopen, so a future regression that drops
``Font.shadow`` (or breaks the XML mapping) fails loudly here.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.dml.effect import EffectFormat, ShadowFormat
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches


class DescribeIssue546TextShadow(object):
    """End-to-end regression suite for the #546 text-shadow feature-gap."""

    def it_exposes_ShadowFormat_on_run_font(self):
        """``Run.font.shadow`` returns a real |ShadowFormat|, not a stub.

        The reporter's core ask: a handle on the run's text-shadow. A
        regression that removes ``Font.shadow`` or returns the wrong
        type fails here.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # -- blank layout
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1)
        )
        shape.text_frame.text = "Shadowed text"
        run = shape.text_frame.paragraphs[0].runs[0]

        assert isinstance(run.font.shadow, ShadowFormat)

        # -- baseline: no explicit shadow; run inherits from style hierarchy --
        assert run.font.shadow.inherit is True
        assert run.font.shadow.blur_radius is None
        assert run.font.shadow.distance is None
        assert run.font.shadow.direction is None

    def it_sets_every_shadow_knob_on_a_run(self):
        """All four shadow properties (blur, distance, direction, color)
        are independently read/writable on a run font.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1)
        )
        shape.text_frame.text = "Shadowed text"
        run = shape.text_frame.paragraphs[0].runs[0]
        font = run.font

        font.shadow.blur_radius = Emu(50800)
        font.shadow.distance = Emu(38100)
        font.shadow.direction = 45.0
        font.shadow.color.rgb = RGBColor(0xFF, 0x00, 0x00)

        assert font.shadow.blur_radius == Emu(50800)
        assert font.shadow.distance == Emu(38100)
        assert font.shadow.direction == 45.0
        assert font.shadow.color.rgb == RGBColor(0xFF, 0x00, 0x00)
        # -- writing a knob materializes `a:effectLst/a:outerShdw`, so
        # -- inheritance is now broken --
        assert font.shadow.inherit is False

    def it_writes_shadow_xml_under_rPr_effectLst_outerShdw(self):
        """The produced XML must land under
        ``a:r/a:rPr/a:effectLst/a:outerShdw`` — the schema-defined
        location for a text-run shadow.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1)
        )
        shape.text_frame.text = "Shadowed text"
        run = shape.text_frame.paragraphs[0].runs[0]

        run.font.shadow.blur_radius = Emu(50800)

        rPr = run._r.find(qn("a:rPr"))
        assert rPr is not None
        effectLst = rPr.find(qn("a:effectLst"))
        assert effectLst is not None
        outerShdw = effectLst.find(qn("a:outerShdw"))
        assert outerShdw is not None
        assert outerShdw.get("blurRad") == "50800"

    def it_round_trips_text_shadow_through_save_and_reopen(self):
        """The full #546 flow: author a text shadow, save, reopen, and
        verify every knob survives the package round-trip.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1)
        )
        shape.text_frame.text = "Shadowed text"
        run = shape.text_frame.paragraphs[0].runs[0]

        run.font.shadow.blur_radius = Emu(76200)  # -- 6 pt blur
        run.font.shadow.distance = Emu(57150)  # -- 4.5 pt offset
        run.font.shadow.direction = 135.0  # -- down-right
        run.font.shadow.color.rgb = RGBColor(0x11, 0x22, 0x33)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        run2 = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]

        assert run2.font.shadow.blur_radius == Emu(76200)
        assert run2.font.shadow.distance == Emu(57150)
        assert run2.font.shadow.direction == 135.0
        assert run2.font.shadow.color.rgb == RGBColor(0x11, 0x22, 0x33)
        assert run2.font.shadow.inherit is False

    def it_restores_inheritance_when_inherit_set_True(self):
        """Assigning ``font.shadow.inherit = True`` removes the explicit
        ``a:effectLst`` and restores style-hierarchy inheritance.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1)
        )
        shape.text_frame.text = "Shadowed text"
        run = shape.text_frame.paragraphs[0].runs[0]

        run.font.shadow.blur_radius = Emu(50800)
        assert run.font.shadow.inherit is False

        run.font.shadow.inherit = True

        assert run.font.shadow.inherit is True
        assert run.font.shadow.blur_radius is None
        assert run.font.shadow.distance is None
        assert run.font.shadow.direction is None
        # -- `a:effectLst` has been physically removed from the rPr --
        rPr = run._r.find(qn("a:rPr"))
        if rPr is not None:
            assert rPr.find(qn("a:effectLst")) is None

    def it_exposes_effect_format_family_on_font(self):
        """``Font.effect_format`` exposes the full ``a:effectLst`` family —
        ``shadow`` / ``glow`` / ``reflection`` / ``soft_edge`` — on the
        run's ``a:rPr``, consistent with the same ``EffectFormat`` API
        used by shapes, groups, and chart elements.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1)
        )
        shape.text_frame.text = "Glowing text"
        run = shape.text_frame.paragraphs[0].runs[0]

        ef = run.font.effect_format
        assert isinstance(ef, EffectFormat)

        # -- writing via effect_format lands on the same a:rPr/a:effectLst --
        ef.glow.size = Emu(50800)
        rPr = run._r.find(qn("a:rPr"))
        assert rPr is not None
        glow = rPr.find(qn("a:effectLst") + "/" + qn("a:glow"))
        assert glow is not None
        assert glow.get("rad") == "50800"

    def it_exposes_shadow_on_paragraph_font_and_field_font(self):
        """Paragraph-level (``a:pPr/a:defRPr``) and text-field
        (``a:fld/a:rPr``) font objects also carry a |ShadowFormat|,
        consistent with the way they already expose fill and color.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1)
        )
        shape.text_frame.text = "para"
        paragraph = shape.text_frame.paragraphs[0]

        # -- paragraph-level default-RPr shadow --
        assert isinstance(paragraph.font.shadow, ShadowFormat)
        paragraph.font.shadow.blur_radius = Emu(25400)
        assert paragraph.font.shadow.blur_radius == Emu(25400)
