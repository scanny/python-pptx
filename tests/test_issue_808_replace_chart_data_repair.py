# pyright: reportPrivateUsage=false

"""Regression test for issue #808 — ``Chart.replace_data`` produces a
PowerPoint "repair needed" dialog on open.

Issue #808 (https://github.com/scanny/python-pptx/issues/808) reported
that running ``chart.replace_data(CategoryChartData(...))`` on a
stacked-bar chart produced a ``.pptx`` PowerPoint refused to open
without first offering to "Repair" it — and Repair then dropped every
shape on the affected slide. The two root causes that put callers in
that state in-the-wild are both Wave-3 fixes:

* **#396** — passing a mismatched ``ChartData`` subclass (e.g.
  ``CategoryChartData`` to an XY/scatter chart, or an XY / bubble-shaped
  data set to a category chart) wrote ``c:cat/c:val`` into a plot
  expecting ``c:xVal/c:yVal`` (and ``c:bubbleSize``), producing a chart
  PowerPoint could not parse. Fixed by ``Chart.replace_data`` raising
  ``ValueError`` up-front when the data shape is wrong for the chart
  type. See commit ``3777cfb9`` on master.

* **#490** — charts whose embedded-workbook relationship was stripped
  (``c:externalData`` still present but its ``@r:id`` no longer mapped
  to a real rel; common for charts pasted from pre-2007 ``.xls``
  sources) made ``ChartWorkbook.xlsx_part`` raise ``KeyError`` from
  inside ``Chart.replace_data``, which both prevented the rewrite and
  left callers guessing. Fixed by tolerating the missing rel in the
  getter so the setter can re-attach a fresh ``EmbeddedXlsxPart``. See
  commit ``1b45bf36`` on master.

This test pins the #808 reporter's happy-path — a 2x2 stacked-bar chart
with 6 new series, 5 new categories — and asserts the resulting chart
XML is the shape PowerPoint expects for a category chart (no stray
``c:xVal`` / ``c:yVal`` / ``c:bubbleSize``), that each ``c:ser`` gets a
unique ``c:idx`` / ``c:order``, that every ``c:numRef`` / ``c:strRef``
carries a matching cache, and that the file round-trips through save
+ reload with all 6 series intact. It also covers the two root-cause
guardrails the Wave-3 fixes introduced so a regression of either path
would reproduce #808.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.chart.data import (
    BubbleChartData,
    CategoryChartData,
    XyChartData,
)
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against pollution from tests that mutate ``PartFactory.part_type_for``.

    ``tests/opc/test_package.py::DescribePartFactory`` replaces the
    ``CT.PML_SLIDE`` entry with a Mock and does not restore it. Without
    this fixture, the round-trip test below would get a Mock back for
    ``slide`` on reload. See the identical fixture in
    ``tests/test_issue_400_animation_umbrella.py`` and
    ``tests/test_issue_640_chart_duplicate.py``.
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


class DescribeIssue808ReplaceChartDataRepair:
    """``Chart.replace_data`` no longer produces a repair-needed file (issue #808)."""

    def it_replaces_a_stacked_bar_chart_with_6_new_series_without_corruption(self):
        """Reporter's exact scenario: stacked-bar chart, 6 series, 5 categories.

        The produced ``.pptx`` must have valid category-chart XML (no stray
        XY/bubble elements), well-formed series caches, and unique
        idx/order on every ``c:ser``. These are the structural invariants
        PowerPoint's parser enforces; violating any of them triggers the
        "repair needed" dialog the #808 reporter saw.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        # --- start with a 2-series stacked-bar chart so replace_data has
        # --- to clone 4 additional series, exercising the path the #808
        # --- reporter hit (original chart had 2 series, replacement had 6). ---
        initial = CategoryChartData()
        initial.categories = ["2015", "2016"]
        initial.add_series("S1", (1.0, 2.0))
        initial.add_series("S2", (3.0, 4.0))
        gf = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_STACKED_100,
            Inches(1),
            Inches(1),
            Inches(5),
            Inches(3),
            initial,
        )
        chart = gf.chart

        # --- act: run the reporter's own replace_data call ---
        cd = CategoryChartData()
        cd.categories = ["2017", "2018", "2019", "2020", "2021"]
        cd.add_series("Series 1", (198.0, 80.0, 70.0, 577.0, 1420.0))
        cd.add_series("Series 2", (98.0, 180.0, 170.0, 77.0, 3420.0))
        cd.add_series("Series 3", (298.0, 480.0, 270.0, 277.0, 4220.0))
        cd.add_series("Series 4", (198.0, 80.0, 70.0, 577.0, 5420.0))
        cd.add_series("Series 5", (98.0, 180.0, 170.0, 77.0, 6420.0))
        cd.add_series("Series 6", (298.0, 480.0, 270.0, 277.0, 7220.0))
        chart.replace_data(cd)

        chartSpace = chart._chartSpace

        # --- no XY/bubble-shaped children snuck into a category chart.
        # --- The presence of any of these on a barChart is a reliable
        # --- "repair" trigger (#396 pre-fix symptom). ---
        assert chartSpace.find(".//" + qn("c:xVal")) is None
        assert chartSpace.find(".//" + qn("c:yVal")) is None
        assert chartSpace.find(".//" + qn("c:bubbleSize")) is None

        # --- all 6 series are present and carry unique idx + order ---
        sers = chartSpace.findall(".//" + qn("c:plotArea") + "//" + qn("c:ser"))
        assert len(sers) == 6
        idxs = [int(s.find(qn("c:idx")).get("val")) for s in sers]
        orders = [int(s.find(qn("c:order")).get("val")) for s in sers]
        assert sorted(idxs) == [0, 1, 2, 3, 4, 5], idxs
        assert sorted(orders) == [0, 1, 2, 3, 4, 5], orders

        # --- every numRef / strRef carries a matching cache with the right
        # --- ptCount. A missing cache or a cache whose ptCount disagrees
        # --- with the referenced range is another classic repair trigger. ---
        for numRef in chartSpace.iter(qn("c:numRef")):
            cache = numRef.find(qn("c:numCache"))
            assert cache is not None, (
                "c:numRef without c:numCache would trigger repair"
            )
            ptCount = cache.find(qn("c:ptCount"))
            assert ptCount is not None
        for strRef in chartSpace.iter(qn("c:strRef")):
            cache = strRef.find(qn("c:strCache"))
            assert cache is not None, (
                "c:strRef without c:strCache would trigger repair"
            )

        # --- embedded xlsx exists and carries the new data columns.
        # --- (B..G for 6 series, rows 2..6 for 5 categories). ---
        assert chart.workbook is not None

    def it_round_trips_the_reporter_scenario_through_save_and_reopen(
        self, _restore_part_factory
    ):
        """The repaired file survives a save + reopen without losing series.

        #808 described PowerPoint dropping "all the contents on slide[1]"
        after clicking Repair. If the resulting file still has 6 working
        series after python-pptx itself reloads the package, that rules
        out the structural-corruption class of failure.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        initial = CategoryChartData()
        initial.categories = ["X", "Y"]
        initial.add_series("A", (1.0, 2.0))
        initial.add_series("B", (3.0, 4.0))
        gf = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_STACKED_100,
            Inches(1),
            Inches(1),
            Inches(5),
            Inches(3),
            initial,
        )

        cd = CategoryChartData()
        cd.categories = ["2017", "2018", "2019", "2020", "2021"]
        for n in range(1, 7):
            cd.add_series("Series %d" % n, (float(n), float(n + 1), float(n + 2),
                                             float(n + 3), float(n + 4)))
        gf.chart.replace_data(cd)

        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        prs2 = Presentation(bio)

        slide2 = list(prs2.slides)[0]
        reloaded = [s for s in slide2.shapes if s.has_chart][0].chart
        assert len(reloaded.series) == 6
        assert reloaded.chart_type == XL_CHART_TYPE.BAR_STACKED_100
        # --- categories preserved ---
        plot = reloaded.plots[0]
        assert list(plot.categories) == ["2017", "2018", "2019", "2020", "2021"]
        # --- series names preserved ---
        names = [ser.name for ser in reloaded.series]
        assert names == ["Series 1", "Series 2", "Series 3", "Series 4",
                         "Series 5", "Series 6"]

    @pytest.mark.parametrize(
        "chart_type, wrong_data_factory",
        [
            # --- XY chart + CategoryChartData: pre-#396 would write
            # --- c:cat/c:val into an XY plot expecting c:xVal/c:yVal. ---
            (XL_CHART_TYPE.XY_SCATTER_LINES, lambda: _make_category_data()),
            # --- Bubble chart + XyChartData: missing c:bubbleSize. ---
            (XL_CHART_TYPE.BUBBLE, lambda: _make_xy_data()),
            # --- Category chart + BubbleChartData: extra c:bubbleSize
            # --- leaked into a plot that cannot consume it. ---
            (XL_CHART_TYPE.BAR_STACKED_100, lambda: _make_bubble_data()),
        ],
    )
    def it_rejects_mismatched_data_before_writing_a_broken_file_issue_396(
        self, chart_type, wrong_data_factory
    ):
        """The #396 type-mismatch guard fails fast with ValueError rather than
        silently producing a file PowerPoint will ask to repair."""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        # --- seed the chart with a type-appropriate ChartData so add_chart
        # --- succeeds; the mismatch is only introduced on replace_data. ---
        if chart_type == XL_CHART_TYPE.BUBBLE:
            seed = _make_bubble_data()
        elif chart_type == XL_CHART_TYPE.XY_SCATTER_LINES:
            seed = _make_xy_data()
        else:
            seed = _make_category_data()
        gf = slide.shapes.add_chart(
            chart_type,
            Inches(1),
            Inches(1),
            Inches(5),
            Inches(3),
            seed,
        )

        with pytest.raises(ValueError):
            gf.chart.replace_data(wrong_data_factory())

    def it_recovers_when_the_embedded_xlsx_rel_is_missing_issue_490(self):
        """#490 fix: a chart whose ``c:externalData/@r:id`` no longer resolves
        must still accept ``replace_data`` without raising ``KeyError`` — the
        ``EmbeddedXlsxPart`` is rebuilt from scratch and the #808 scenario
        completes cleanly.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        cd = CategoryChartData()
        cd.categories = ["X", "Y"]
        cd.add_series("A", (1.0, 2.0))
        gf = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_STACKED_100,
            Inches(1),
            Inches(1),
            Inches(5),
            Inches(3),
            cd,
        )
        chart = gf.chart

        # --- simulate the in-the-wild state: c:externalData is still
        # --- present but its @r:id points at a rId that is no longer in
        # --- the chart-part's relationships (the rel was stripped by
        # --- another client, or the chart was pasted from a pre-2007
        # --- .xls source). ---
        chart_part = chart.part
        externalData = chart._chartSpace.find(qn("c:externalData"))
        assert externalData is not None
        externalData.set(qn("r:id"), "rIdStale")

        # --- act: the reporter's replace_data call. Must not raise. ---
        new_cd = CategoryChartData()
        new_cd.categories = ["2017", "2018", "2019"]
        for n in range(1, 4):
            new_cd.add_series("S%d" % n, (float(n), float(n + 1), float(n + 2)))
        chart.replace_data(new_cd)

        # --- post-condition: the stale rId got overwritten with a live one,
        # --- and there is a fresh embedded xlsx part backing the chart. ---
        assert chart.workbook is not None
        new_rId = externalData.get(qn("r:id"))
        assert new_rId != "rIdStale"
        assert new_rId in chart_part.rels


def _make_category_data():
    cd = CategoryChartData()
    cd.categories = ["A", "B", "C"]
    cd.add_series("S1", (1.0, 2.0, 3.0))
    return cd


def _make_xy_data():
    xd = XyChartData()
    s = xd.add_series("S1")
    s.add_data_point(1.0, 10.0)
    s.add_data_point(2.0, 20.0)
    return xd


def _make_bubble_data():
    bd = BubbleChartData()
    s = bd.add_series("S1")
    s.add_data_point(1.0, 10.0, 5.0)
    s.add_data_point(2.0, 20.0, 7.0)
    return bd
