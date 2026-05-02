# pyright: reportPrivateUsage=false

"""Regression test for issue #705 — shape shadows.

Issue #705 (https://github.com/scanny/python-pptx/issues/705) asks:
"status of the project? support for shape shadows?"

The original ``BaseShape.shadow`` shipped as a skeletal ``ShadowFormat``
with only an ``.inherit`` boolean — the four shadow knobs the reporter
actually needed (blur radius, offset distance, offset direction, color)
were not exposed. Two subsequent changes deliver the full feature:

  * Foundation ``F2`` (Wave 1) added the complete ``a:effectLst`` family
    (``a:outerShdw`` / ``a:glow`` / ``a:reflection`` / ``a:softEdge``)
    with read/write ``ShadowFormat.blur_radius`` / ``distance`` /
    ``direction`` / ``color`` properties wired through
    ``pptx.dml.effect.ShadowFormat`` (``src/pptx/dml/effect.py``).
  * Issue #130 (Wave 3) lifted the same ``ShadowFormat`` onto
    ``ChartFormat.shadow`` so a chart plot-area / chart-space can carry
    a shadow alongside ordinary auto-shapes.
  * Issue #446 (Wave 2) tightened the inherit semantics: setting
    ``shadow.inherit = False`` now also zeroes any sibling
    ``p:style/a:effectRef/@idx`` so the theme-level shadow inherited
    via ``effectRef`` doesn't leak through the explicit (empty)
    ``a:effectLst``.

This regression test pins the full API surface the #705 reporter asked
for and round-trips every knob through ``Presentation.save`` + reopen
so a future refactor that drops any of the four properties (or breaks
their XML mapping) fails loudly here rather than at user-report time.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.dml.effect import ShadowFormat
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches


class DescribeIssue705ShapeShadows(object):
    """End-to-end regression suite for the #705 shape-shadow feature-gap."""

    def it_exposes_the_full_ShadowFormat_api_on_shape_shadow(self):
        """Pin the exact property names / types the #705 reporter asked for.

        A regression that removes any of ``blur_radius``, ``distance``,
        ``direction``, or ``color`` — or that renames them / moves them
        off of ``BaseShape.shadow`` — fails this test immediately.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # -- blank layout
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
        )

        # -- `.shadow` is a real ShadowFormat, not the skeletal stub --
        assert isinstance(shape.shadow, ShadowFormat)

        # -- baseline: nothing configured, all four knobs are None and
        # -- the shape inherits its shadow from the theme --
        assert shape.shadow.inherit is True
        assert shape.shadow.blur_radius is None
        assert shape.shadow.distance is None
        assert shape.shadow.direction is None

        # -- each property is independently read/writable --
        shape.shadow.blur_radius = Emu(50800)
        shape.shadow.distance = Emu(38100)
        shape.shadow.direction = 270.0
        shape.shadow.color.rgb = RGBColor(0xFF, 0x00, 0x00)

        assert shape.shadow.blur_radius == Emu(50800)
        assert shape.shadow.distance == Emu(38100)
        assert shape.shadow.direction == 270.0
        assert shape.shadow.color.rgb == RGBColor(0xFF, 0x00, 0x00)

        # -- writing a knob materializes `a:effectLst` / `a:outerShdw`, so
        # -- inheritance is now broken --
        assert shape.shadow.inherit is False

    def it_round_trips_shadow_settings_through_save_and_reopen(
        self
    ):
        """The #705 flow end-to-end: author a shadow, save, reopen, assert
        every knob survived.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1.5)
        )

        shape.shadow.blur_radius = Emu(76200)  # -- 6 pt blur
        shape.shadow.distance = Emu(57150)  # -- 4.5 pt offset
        shape.shadow.direction = 135.0  # -- down-right
        shape.shadow.color.rgb = RGBColor(0x11, 0x22, 0x33)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        shape2 = prs2.slides[0].shapes[0]

        # -- every shadow property must survive the round-trip --
        assert shape2.shadow.blur_radius == Emu(76200)
        assert shape2.shadow.distance == Emu(57150)
        assert shape2.shadow.direction == 135.0
        assert shape2.shadow.color.rgb == RGBColor(0x11, 0x22, 0x33)
        assert shape2.shadow.inherit is False

    def it_restores_inheritance_when_inherit_is_set_True(self):
        """Assigning ``shadow.inherit = True`` removes the explicit
        ``a:effectLst`` and restores theme inheritance — the symmetric
        behaviour #446 locked down.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
        )

        shape.shadow.blur_radius = Emu(50800)
        assert shape.shadow.inherit is False

        shape.shadow.inherit = True

        assert shape.shadow.inherit is True
        assert shape.shadow.blur_radius is None
        assert shape.shadow.distance is None
        assert shape.shadow.direction is None

    def it_zeroes_sibling_effectRef_idx_when_breaking_inheritance(self):
        """Issue #446 semantics: setting ``inherit = False`` on a shape
        that carries a theme-level shadow via ``p:style/a:effectRef@idx``
        must also zero that idx, so the explicit (empty) ``a:effectLst``
        really does suppress the inherited shadow.
        """
        from pptx.oxml.ns import qn

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
        )

        # -- `add_shape` emits a `p:style/a:effectRef` with a non-zero idx --
        sp = shape._element
        effectRef = sp.find(f"./{qn('p:style')}/{qn('a:effectRef')}")
        assert effectRef is not None
        assert effectRef.get("idx") != "0"

        shape.shadow.inherit = False

        # -- effectRef idx is zeroed (issue #446) --
        assert effectRef.get("idx") == "0"

    def it_exposes_shadow_on_group_shapes_too(self):
        """#705 tagged both ``shapes/base.py`` and ``shapes/group.py``.
        Verify ``GroupShape.shadow`` is also a real |ShadowFormat| with
        the blur / distance / direction / color API (not just a stub).
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # -- add two rectangles, group them --
        r1 = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1)
        )
        r2 = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(3), Inches(1), Inches(1), Inches(1)
        )
        group = slide.shapes.add_group_shape([r1, r2])

        assert isinstance(group.shadow, ShadowFormat)
        assert group.shadow.inherit is True

        group.shadow.blur_radius = Emu(25400)
        group.shadow.distance = Emu(19050)
        group.shadow.direction = 90.0

        assert group.shadow.blur_radius == Emu(25400)
        assert group.shadow.distance == Emu(19050)
        assert group.shadow.direction == 90.0
        assert group.shadow.inherit is False
