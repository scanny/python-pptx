# pyright: reportPrivateUsage=false

"""Regression test for issue #275 — controlling shadows on autoshapes.

Issue #275 (https://github.com/scanny/python-pptx/issues/275) asks:

    "I am trying to create some presentations programatically.
     I am using autoshape from pptx0.6.
     Do we have any option to control the shadow of the shape inserted
     using MSO_SHAPE"

At the time of the report, ``BaseShape.shadow`` was a skeletal
``ShadowFormat`` exposing only an ``.inherit`` boolean — the knobs a
caller needs to *control* an autoshape shadow (blur radius, offset
distance, offset direction, color) were not exposed.

Three subsequent changes ship the full feature the #275 reporter
asked for:

  * Foundation ``F2`` (Wave 1) added the complete ``a:effectLst`` family
    (``a:outerShdw`` / ``a:glow`` / ``a:reflection`` / ``a:softEdge``)
    with read/write ``ShadowFormat.blur_radius`` / ``distance`` /
    ``direction`` / ``color`` properties wired through
    ``pptx.dml.effect.ShadowFormat`` (``src/pptx/dml/effect.py``).
  * Issue #446 (Wave 2) tightened the ``inherit = False`` semantics so
    that zeroing inheritance also clears any sibling
    ``p:style/a:effectRef/@idx``, preventing theme-level shadow from
    leaking through an explicit (empty) ``a:effectLst``.
  * Issue #130 (Wave 3) lifted the same full-knob ``ShadowFormat`` onto
    ``ChartFormat.shadow`` so charts participate in the same API.
  * Issue #705 (Wave 7) pinned the API end-to-end on ``Shape.shadow`` /
    ``GroupShape.shadow`` with a regression suite.

Issue #275 predates #705 and is the same underlying feature-gap
restricted to the specific ``slide.shapes.add_shape(MSO_SHAPE.…)``
call path the #275 reporter names. This test pins the #275 reporter's
exact workflow — create an autoshape with ``MSO_SHAPE``, set each of
the four shadow knobs, round-trip through ``Presentation.save`` plus
reopen, and assert every knob survived — so a future regression of
any of the three fixes above fails loudly against the original
report's wording.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.dml.effect import ShadowFormat
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches


class DescribeIssue275AutoshapeShadows(object):
    """End-to-end regression suite for the #275 feature request."""

    def it_controls_shadow_on_an_MSO_SHAPE_autoshape(self):
        """The reporter's exact use case: an ``MSO_SHAPE`` autoshape plus
        a caller-chosen shadow. All four knobs must round-trip via the
        in-memory ``ShadowFormat`` API on the freshly-added shape.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # -- blank layout

        # -- reporter's code path: `add_shape(MSO_SHAPE.…)` --
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1.5)
        )

        # -- `.shadow` is the full ShadowFormat, not the pre-F2 stub --
        assert isinstance(shape.shadow, ShadowFormat)

        # -- baseline: no explicit shadow, shape inherits from theme --
        assert shape.shadow.inherit is True
        assert shape.shadow.blur_radius is None
        assert shape.shadow.distance is None
        assert shape.shadow.direction is None

        # -- the four knobs the reporter needs to "control the shadow" --
        shape.shadow.blur_radius = Emu(63500)  # -- 5 pt blur
        shape.shadow.distance = Emu(50800)  # -- 4 pt offset
        shape.shadow.direction = 45.0  # -- up-right
        shape.shadow.color.rgb = RGBColor(0x80, 0x80, 0x80)  # -- mid-gray

        assert shape.shadow.blur_radius == Emu(63500)
        assert shape.shadow.distance == Emu(50800)
        assert shape.shadow.direction == 45.0
        assert shape.shadow.color.rgb == RGBColor(0x80, 0x80, 0x80)

        # -- configuring any knob materializes `a:effectLst / a:outerShdw`,
        # -- so the shape no longer inherits its shadow from the theme --
        assert shape.shadow.inherit is False

    def it_round_trips_autoshape_shadow_through_save_and_reopen(self):
        """Full #275 flow: author an autoshape shadow, save the deck,
        reopen, and assert every knob survived the XML round-trip.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(2), Inches(2)
        )

        shape.shadow.blur_radius = Emu(88900)  # -- 7 pt blur
        shape.shadow.distance = Emu(44450)  # -- 3.5 pt offset
        shape.shadow.direction = 225.0  # -- down-left
        shape.shadow.color.rgb = RGBColor(0xAA, 0xBB, 0xCC)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        shape2 = prs2.slides[0].shapes[0]

        assert shape2.shadow.blur_radius == Emu(88900)
        assert shape2.shadow.distance == Emu(44450)
        assert shape2.shadow.direction == 225.0
        assert shape2.shadow.color.rgb == RGBColor(0xAA, 0xBB, 0xCC)
        assert shape2.shadow.inherit is False

    def it_suppresses_theme_shadow_when_inherit_is_set_False(self):
        """Assigning ``shadow.inherit = False`` on a newly-added autoshape
        must also zero the sibling ``p:style/a:effectRef/@idx`` (issue
        #446) so the empty ``a:effectLst`` really does suppress the
        theme-level shadow a user might otherwise see on an MSO_SHAPE.
        """
        from pptx.oxml.ns import qn

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
        )

        # -- `add_shape` emits `p:style/a:effectRef` with a non-zero idx --
        sp = shape._element
        effectRef = sp.find(f"./{qn('p:style')}/{qn('a:effectRef')}")
        assert effectRef is not None
        assert effectRef.get("idx") != "0"

        shape.shadow.inherit = False

        # -- and after breaking inheritance, idx is zeroed (#446) --
        assert effectRef.get("idx") == "0"
