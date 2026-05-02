# pyright: reportPrivateUsage=false

"""Regression test for issue #539 — bar chart color duplicated when extending series.

Issue #539 (https://github.com/scanny/python-pptx/issues/539) reports that
after calling ``Chart.replace_data()`` to grow a chart from N series to
M > N series, all the newly-added series are rendered in the same colour
as the last source series — PowerPoint shows every extended column in
the source's colour rather than cycling through the theme accents.

Wave 2 resolved this under ``feat/issue-529-chart-theme-colors``
(commit ``2e078d7e``) by introducing
``pptx.chart.xmlwriter._apply_accent_color_to_ser`` and calling it from
``_BaseSeriesXmlRewriter._add_cloned_sers``. Every cloned ``c:ser`` now
gets its explicit ``a:srgbClr`` / ``a:schemeClr`` fills rewritten to
``a:schemeClr val="accent{n}"`` where ``n`` cycles 1..6 on the new
series's index.

This regression test reproduces the exact flow the #539 reporter posted
(paint the source series red, grow from 1 to 6 series via
``replace_data``) and pins the wave-2 resolution so any future change
that reintroduces the "all new series repeat the source colour" bug
will be caught here.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches


def _chart_of(prs):
    """Return the first chart on the first slide of *prs*."""
    slide = list(prs.slides)[0]
    for shp in slide.shapes:
        if shp.has_chart:
            return shp.chart
    raise AssertionError("no chart found on first slide")


class DescribeIssue539BarChartColorCycle:
    """Growing a colored 1-series chart cycles theme accents on the new series."""

    def it_does_not_duplicate_source_color_on_new_series(self):
        """The #539 reporter's exact scenario.

        A 1-series COLUMN_CLUSTERED chart painted red, grown to 6 series
        via ``replace_data``, produces an ser[0] that keeps the explicit
        red and ser[1..5] that carry ``a:schemeClr val="accent{2..6}"``
        instead of duplicating the red fill onto every clone.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        cd = CategoryChartData()
        cd.categories = ["Italy", "France", "Spain"]
        cd.add_series("Model 1", (19.2, 21.4, 16.7))
        gf = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            cd,
        )
        chart = gf.chart

        # -- paint the only source series red (the #539 precondition) --
        source_series = chart.series[0]
        source_series.format.fill.solid()
        source_series.format.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)

        # -- act: grow to 6 series via replace_data (the #539 flow) --
        cd2 = CategoryChartData()
        cd2.categories = ["Italy", "France", "Spain", "Germany", "Holland", "Belgium"]
        cd2.add_series("Model 1", (19.2, 21.4, 16.7, 15, 19.3, 11))
        cd2.add_series("Model 2", (22.3, 28.6, 15.2, 15, 15, 10))
        cd2.add_series("Model 3", (20.4, 26.3, 14.2, 15, 22, 13))
        cd2.add_series("Model 4", (21.4, 22.3, 18.2, 15, 25, 18))
        cd2.add_series("Model 5", (21.4, 18, 18.2, 10, 25, 14))
        cd2.add_series("Model 6", (21.4, 18, 18.2, 10, 25, 14))
        chart.replace_data(cd2)

        # -- inspect every ser --
        sers = chart._chartSpace.xpath(".//c:ser")
        assert len(sers) == 6

        # -- ser[0] (the original red source) keeps its explicit sRGB fill --
        assert sers[0].xpath(".//a:srgbClr/@val") == ["FF0000"]
        assert sers[0].xpath(".//a:schemeClr") == []

        # -- ser[1..5] (the five clones) must each (a) NOT carry the red
        # -- source colour and (b) carry exactly one schemeClr pointing at
        # -- the next accent in the cycle. This is the #539 bug: before the
        # -- fix, every clone would have a:srgbClr val="FF0000" here.
        expected_accents = ["accent2", "accent3", "accent4", "accent5", "accent6"]
        for i, accent in zip(range(1, 6), expected_accents):
            assert sers[i].xpath(".//a:srgbClr") == [], (
                "ser[%d] still carries an explicit sRGB fill — #539 regressed" % i
            )
            assert sers[i].xpath(".//a:schemeClr/@val") == [accent], (
                "ser[%d] has accents %r, expected [%r]"
                % (i, sers[i].xpath(".//a:schemeClr/@val"), accent)
            )

    def it_cycles_accents_past_six_wrapping_to_accent1(self):
        """Extending to more than 6 series wraps the accent cycle.

        ser[6] should land on accent1, ser[7] on accent2, etc.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        cd = CategoryChartData()
        cd.categories = ["A", "B", "C"]
        cd.add_series("S1", (1.0, 2.0, 3.0))
        gf = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(5),
            Inches(3),
            cd,
        )
        chart = gf.chart
        # -- paint source red so the clone path goes through the accent rewrite --
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)

        # -- grow to 8 series so the cycle wraps --
        cd2 = CategoryChartData()
        cd2.categories = ["A", "B", "C"]
        for i in range(8):
            cd2.add_series("S%d" % (i + 1), (1.0, 2.0, 3.0))
        chart.replace_data(cd2)

        sers = chart._chartSpace.xpath(".//c:ser")
        assert len(sers) == 8

        # -- the six clones at idx 1..6 cycle accent2..accent6 then wrap to accent1 --
        expected = [
            "accent2",  # idx 1
            "accent3",  # idx 2
            "accent4",  # idx 3
            "accent5",  # idx 4
            "accent6",  # idx 5
            "accent1",  # idx 6 — wrap
            "accent2",  # idx 7 — wrap continues
        ]
        for i, accent in zip(range(1, 8), expected):
            assert sers[i].xpath(".//a:schemeClr/@val") == [accent], (
                "ser[%d] got %r, expected %r"
                % (i, sers[i].xpath(".//a:schemeClr/@val"), accent)
            )

    def it_survives_save_and_reopen_without_reverting_to_source_color(
        self
    ):
        """The fix must survive a full save + Presentation() round-trip.

        The cloned sers must still carry their schemeClr accents, not
        revert to the red explicit fill, after the file is saved and
        reopened.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        cd = CategoryChartData()
        cd.categories = ["Q1", "Q2", "Q3"]
        cd.add_series("Revenue", (10.0, 20.0, 30.0))
        gf = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            cd,
        )
        chart = gf.chart
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)

        cd2 = CategoryChartData()
        cd2.categories = ["Q1", "Q2", "Q3"]
        for i in range(6):
            cd2.add_series("Series %d" % (i + 1), (1.0, 2.0, 3.0))
        chart.replace_data(cd2)

        # -- round-trip --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        chart2 = _chart_of(prs2)

        sers = chart2._chartSpace.xpath(".//c:ser")
        assert len(sers) == 6

        # -- ser[0] (the source) keeps its red --
        assert sers[0].xpath(".//a:srgbClr/@val") == ["FF0000"]

        # -- ser[1..5] must still NOT carry FF0000 after reload --
        for i in range(1, 6):
            assert "FF0000" not in sers[i].xpath(".//a:srgbClr/@val"), (
                "ser[%d] reverted to source red after reload — #539 regressed" % i
            )

        # -- and they still carry the expected accent sequence --
        expected = ["accent2", "accent3", "accent4", "accent5", "accent6"]
        for i, accent in zip(range(1, 6), expected):
            assert sers[i].xpath(".//a:schemeClr/@val") == [accent]
