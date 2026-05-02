# pyright: reportPrivateUsage=false

"""Regression test for issue #1030 — manually set chart title position.

Issue #1030 (https://github.com/scanny/python-pptx/issues/1030) asks for a
way to manually position a chart title. PowerPoint's UI lets a user drag
the title to an arbitrary location over the chart, persisting the result
in the chart XML as ``c:title/c:layout/c:manualLayout`` with ``c:x`` and
``c:y`` children in the 0.0-1.0 relative-to-chart coordinate space.

This test pins the resolution: ``ChartTitle.position`` is a read/write
property returning a ``(x, y)`` tuple when the title carries a full
``c:manualLayout`` with ``xMode``/``yMode`` set to ``factor``, or ``None``
when the title is laid out automatically. Assigning a tuple writes the
full ``c:layout/c:manualLayout/(c:xMode,c:yMode,c:x,c:y)`` subtree;
assigning ``None`` removes ``c:layout`` entirely.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against pollution from other tests that mutate ``PartFactory.part_type_for``."""
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


def _add_chart_with_title(prs):
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
    chart.has_title = True
    chart.chart_title.text_frame.text = "Sales"
    return chart


class DescribeIssue1030ChartTitlePosition:
    """``ChartTitle.position`` reads/writes ``c:title/c:layout/c:manualLayout``."""

    def it_returns_None_for_an_auto_laid_out_title(self, _restore_part_factory):
        prs = Presentation()
        chart = _add_chart_with_title(prs)
        # -- a freshly-added title has only <c:layout/> with no manualLayout --
        assert chart.chart_title.position is None

    def it_writes_manual_layout_when_position_is_assigned(self, _restore_part_factory):
        prs = Presentation()
        chart = _add_chart_with_title(prs)

        chart.chart_title.position = (0.25, 0.5)

        assert chart.chart_title.position == (0.25, 0.5)
        title = chart._chartSpace.xpath(".//c:title")[0]
        x_vals = title.xpath("c:layout/c:manualLayout/c:x/@val")
        y_vals = title.xpath("c:layout/c:manualLayout/c:y/@val")
        assert x_vals == ["0.25"]
        assert y_vals == ["0.5"]
        # -- both xMode and yMode must be present; they default to "factor"
        # -- so val may be absent, but the element is what drives
        # -- horz_offset-style readers --
        assert title.xpath("c:layout/c:manualLayout/c:xMode")
        assert title.xpath("c:layout/c:manualLayout/c:yMode")

    def it_removes_layout_when_position_set_to_None(self, _restore_part_factory):
        prs = Presentation()
        chart = _add_chart_with_title(prs)
        chart.chart_title.position = (0.3, 0.4)
        assert chart.chart_title.position == (0.3, 0.4)

        chart.chart_title.position = None

        assert chart.chart_title.position is None
        title = chart._chartSpace.xpath(".//c:title")[0]
        assert title.xpath("c:layout") == []

    def it_round_trips_position_across_save_and_reopen(self, _restore_part_factory):
        prs = Presentation()
        chart = _add_chart_with_title(prs)
        chart.chart_title.position = (0.1, 0.9)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        slide2 = list(prs2.slides)[-1]
        chart2 = next(shp.chart for shp in slide2.shapes if shp.has_chart)
        assert chart2.chart_title.position == (0.1, 0.9)

    def it_rejects_non_tuple_position_values(self, _restore_part_factory):
        prs = Presentation()
        chart = _add_chart_with_title(prs)
        with pytest.raises(ValueError):
            chart.chart_title.position = 0.5
        with pytest.raises(ValueError):
            chart.chart_title.position = (0.1, 0.2, 0.3)
