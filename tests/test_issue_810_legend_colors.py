# pyright: reportPrivateUsage=false

"""Regression test for issue #810 — chart legend colours not matching the graph.

Issue #810 (https://github.com/scanny/python-pptx/issues/810) reports that
the colours shown in a chart's legend do not match the colours of the
corresponding bars / columns / lines in the plot area. The issue was
filed with a vague description (no repro code or sample file) and later
answered by :commit:`2e078d7e` (``feat/issue-529-chart-theme-colors``):
the root cause is that before the fix ``Chart.replace_data()`` cloned
the last source series including its explicit ``a:srgbClr`` fill onto
every new series, so PowerPoint rendered every cloned bar in the source
colour while the legend (which is also drawn from each ``c:ser``'s
``c:spPr``) would pick up whatever accent the theme assigned by series
index for the idx values that had no explicit override — producing the
"legend doesn't match the bars" symptom the reporter described.

Wave 2 resolved this by introducing
``pptx.chart.xmlwriter._apply_accent_color_to_ser``, which on every
cloned ``c:ser`` rewrites descendant ``a:srgbClr`` / ``a:schemeClr``
colours to ``a:schemeClr val="accent{n}"`` where ``n`` cycles 1..6 on
the new series index — matching PowerPoint's own theme-accent rotation.
Because the legend swatch and the plot-area bar for a given series are
both rendered from that series's ``c:spPr``, rewriting them to a single
consistent ``schemeClr`` per series means the legend colour and the bar
colour are guaranteed to match.

This regression test pins that invariant by:

1. Reproducing the #810 scenario the #529 fix addressed (paint the
   source series red, grow to 6 series via ``replace_data``) and
   asserting the plot-area fill on every series equals the legend
   swatch's fill for that series — since both are literally the same
   XML node (``c:ser/c:spPr``), that equivalence is what the fix
   enforces.
2. Round-tripping the same chart through ``Presentation.save`` +
   ``Presentation(buf)`` to confirm the legend-matches-bars invariant
   survives serialization.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against pollution from other tests that mutate ``PartFactory.part_type_for``.

    See the identical fixture in ``tests/test_issue_539_chart_color_cycle.py``
    for the rationale.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.comments import CommentAuthorsPart, CommentsPart
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.parts.image import ImagePart
    from pptx.parts.media import MediaPart
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.slide import (
        NotesMasterPart,
        NotesSlidePart,
        SlideLayoutPart,
        SlideMasterPart,
        SlidePart,
    )

    saved = dict(PartFactory.part_type_for)
    expected = {
        CT.PML_PRESENTATION_MAIN: PresentationPart,
        CT.PML_PRES_MACRO_MAIN: PresentationPart,
        CT.PML_TEMPLATE_MAIN: PresentationPart,
        CT.PML_SLIDESHOW_MAIN: PresentationPart,
        CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
        CT.PML_COMMENTS: CommentsPart,
        CT.PML_COMMENT_AUTHORS: CommentAuthorsPart,
        CT.PML_NOTES_MASTER: NotesMasterPart,
        CT.PML_NOTES_SLIDE: NotesSlidePart,
        CT.PML_SLIDE: SlidePart,
        CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
        CT.PML_SLIDE_MASTER: SlideMasterPart,
        CT.DML_CHART: ChartPart,
        CT.JPEG: ImagePart,
        CT.PNG: ImagePart,
        CT.MP4: MediaPart,
    }
    PartFactory.part_type_for.update(expected)
    try:
        yield
    finally:
        PartFactory.part_type_for.clear()
        PartFactory.part_type_for.update(saved)


def _chart_of(prs):
    """Return the first chart on the first slide of *prs*."""
    slide = list(prs.slides)[0]
    for shp in slide.shapes:
        if shp.has_chart:
            return shp.chart
    raise AssertionError("no chart found on first slide")


def _series_color_key(ser):
    """Return a ``(kind, val)`` tuple identifying the series fill color.

    Used as a comparison key for "does ser[i] render in the same colour
    as the legend swatch for ser[i]?". Since PowerPoint draws both the
    plot-area element (bar, column, slice) and the legend swatch for a
    given series from the same ``c:ser/c:spPr/a:solidFill/*`` node, this
    function just extracts whatever is there and returns it in a form
    that can be compared across series — a series with ``a:srgbClr
    val="FF0000"`` returns ``("srgb", "FF0000")``, a series with
    ``a:schemeClr val="accent2"`` returns ``("scheme", "accent2")``,
    and a series without any explicit colour returns ``("unset",
    None)`` (meaning PowerPoint will pick the theme default by idx).
    """
    srgbs = ser.xpath(".//c:spPr//a:srgbClr/@val")
    if srgbs:
        return ("srgb", srgbs[0])
    schemes = ser.xpath(".//c:spPr//a:schemeClr/@val")
    if schemes:
        return ("scheme", schemes[0])
    return ("unset", None)


class DescribeIssue810LegendColorsMatchGraph:
    """Legend marker colours match plot-area colours across every series."""

    def it_assigns_distinct_colors_to_every_series_after_replace_data(self):
        """Every series ends up with a distinct, explicit colour.

        The #810 root cause was that ``replace_data`` cloned the source
        series's colour onto every new series, so a six-series chart
        would show only two colour groups (source-red plus the accents
        the theme happened to provide for the idxs with no override) —
        the "legend doesn't match the graph" symptom. After the #529
        fix, every clone gets a unique accent, so the six series carry
        six distinct colour keys.
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
        chart.has_legend = True
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)

        cd2 = CategoryChartData()
        cd2.categories = ["Q1", "Q2", "Q3"]
        for i in range(6):
            cd2.add_series("Series %d" % (i + 1), (1.0, 2.0, 3.0))
        chart.replace_data(cd2)

        sers = chart._chartSpace.xpath(".//c:ser")
        assert len(sers) == 6

        # -- every series has a distinct colour key. Before the #529 fix
        # -- all six clones would share ``("srgb", "FF0000")``; after
        # -- the fix ser[0] keeps FF0000 and ser[1..5] get schemeClr
        # -- accent2..accent6 — six distinct keys.
        keys = [_series_color_key(s) for s in sers]
        assert len(set(keys)) == 6, "legend/graph colours collide — #810 regressed: got %r" % (
            keys,
        )

    def it_renders_legend_and_plot_area_from_the_same_series_spPr(self):
        """Legend swatch and plot-area element share one ``c:ser/c:spPr``.

        PowerPoint draws the legend marker for series ``i`` and the
        corresponding plot-area bar from the same ``c:ser/c:spPr`` fill.
        Python-pptx never writes a separate ``c:legendEntry/c:txPr``
        colour, nor a separate plot-area-only colour. So after the #529
        fix, the colour that drives the legend swatch *is* the colour
        that drives the bar for the same series — by construction. This
        test pins that structural guarantee: for each series, the
        single ``c:spPr`` subtree under ``c:ser`` is the only fill-
        colour source, and ``c:legendEntry`` overrides are absent.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        cd = CategoryChartData()
        cd.categories = ["A", "B", "C"]
        cd.add_series("S1", (1.0, 2.0, 3.0))
        gf = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            cd,
        )
        chart = gf.chart
        chart.has_legend = True
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)

        cd2 = CategoryChartData()
        cd2.categories = ["A", "B", "C"]
        for i in range(4):
            cd2.add_series("S%d" % (i + 1), (1.0, 2.0, 3.0))
        chart.replace_data(cd2)

        chartSpace = chart._chartSpace
        # -- there should be no ``c:legendEntry`` colour overrides; if there
        # -- were, legend markers could diverge from plot-area fills.
        legend_entry_fills = chartSpace.xpath(".//c:legend//c:legendEntry//a:solidFill")
        assert (
            legend_entry_fills == []
        ), "c:legendEntry fill override would desynchronize legend from graph: %r" % (
            legend_entry_fills,
        )

        # -- each series has exactly one ``c:spPr/a:solidFill`` (the
        # -- single source of truth feeding both the plot-area fill and
        # -- the legend marker for that series).
        for i, ser in enumerate(chartSpace.xpath(".//c:ser")):
            fills = ser.xpath("./c:spPr/a:solidFill")
            assert len(fills) == 1, (
                "ser[%d] has %d c:spPr/a:solidFill fills, expected 1 (single"
                " source of truth for legend + graph)" % (i, len(fills))
            )

    def it_keeps_legend_matched_to_graph_after_save_and_reopen(self, _restore_part_factory):
        """Round-trip preserves the one-colour-per-series invariant.

        The #810 complaint ("legend not showing same colors as graph")
        is only useful if the file the user actually opens in
        PowerPoint still has the matching colours. This test saves the
        presentation and re-opens it, then reverifies every series
        colour key is distinct and its ``c:ser/c:spPr`` is intact.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        cd = CategoryChartData()
        cd.categories = ["X", "Y", "Z"]
        cd.add_series("One", (1.0, 2.0, 3.0))
        gf = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(6),
            Inches(4),
            cd,
        )
        chart = gf.chart
        chart.has_legend = True
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)

        cd2 = CategoryChartData()
        cd2.categories = ["X", "Y", "Z"]
        for i in range(6):
            cd2.add_series("S%d" % (i + 1), (1.0, 2.0, 3.0))
        chart.replace_data(cd2)

        # -- round-trip --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        chart2 = _chart_of(prs2)

        sers = chart2._chartSpace.xpath(".//c:ser")
        assert len(sers) == 6

        keys = [_series_color_key(s) for s in sers]
        # -- ser[0] still FF0000, ser[1..5] still distinct accents --
        assert keys[0] == ("srgb", "FF0000")
        assert len(set(keys)) == 6, "after save+reopen, legend/graph colours collide — got %r" % (
            keys,
        )
        # -- and no c:legendEntry fill override was introduced during
        # -- serialization that would let the legend diverge --
        assert chart2._chartSpace.xpath(".//c:legend//c:legendEntry//a:solidFill") == []
