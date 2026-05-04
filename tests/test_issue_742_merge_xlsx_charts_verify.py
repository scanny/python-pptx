# pyright: reportPrivateUsage=false

"""Regression test for issue #742 — merging pptx files with Excel graphs/charts.

Issue #742 (https://github.com/scanny/python-pptx/issues/742) reports that
merging decks whose slides carry charts backed by *distinct* embedded
``.xlsx`` workbooks produces a merged deck in which chart data is lost or
cross-pollinated — i.e. more than one chart ends up pointing at the same
workbook, the merged chart's "Edit Data" dialog surfaces values that
belonged to another chart, or the chart parts fail to materialise in the
target package at all.

This scenario is served end-to-end by the slide-copy wave:

  * Wave 1 #1036 (``Slides.add_slide_from_external``) introduced the
    cross-package slide cloner.
  * Wave 7 #934 (``Presentation.merge``) promoted it to the bulk merge
    path and shipped the F5 ``clone_embedded_xlsx`` helper that
    materialises a *fresh* :class:`EmbeddedXlsxPart` for every chart the
    cloner copies — so no two charts in the target deck share a workbook
    with each other or with the source deck.
  * Wave 15 #835 added a regression suite (
    ``tests/test_issue_835_merge_charts_verify.py``) that pins the
    "chart slides survive merge" half of the guarantee.

This verify suite closes out #742 by pinning the *embedded-workbook
isolation* half: after :meth:`Presentation.merge`, every chart in the
target deck carries its own distinct ``.xlsx`` blob whose byte contents
differ from every other chart's blob, and the values PowerPoint would
surface from each merged chart's "Edit Data" dialog match the source
values exactly — both before and after a save+reopen round-trip.

Cross-references:
  * Resolver commits: #934 (``Presentation.merge``), #1036
    (``Slides.add_slide_from_external``).
  * Companion regression: #835
    (``tests/test_issue_835_merge_charts_verify.py``) — chart-carrying
    slide copy, series-values round-trip.
"""

from __future__ import annotations

import io
from typing import TYPE_CHECKING, cast

import pytest

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.parts.chart import ChartPart
from pptx.parts.embeddedpackage import EmbeddedXlsxPart
from pptx.util import Inches

if TYPE_CHECKING:
    from pptx.chart.chart import Chart
    from pptx.presentation import Presentation as PresentationClass
    from pptx.slide import Slide


# -- three distinct chart fingerprints, each distinguishable by categories,
# -- values, AND series name so every chart's embedded xlsx is a different
# -- blob of bytes ----------------------------------------------------------
_CHART_DATA = [
    (("Jan", "Feb", "Mar"), (11.0, 22.0, 33.0), "Revenue"),
    (("Jan", "Feb", "Mar"), (44.0, 55.0, 66.0), "Expense"),
    (("Jan", "Feb", "Mar"), (77.0, 88.0, 99.0), "Profit"),
]


@pytest.fixture
def _restore_part_factory():
    """Guard against ``PartFactory.part_type_for`` pollution from other tests.

    Mirrors the identical guard used by the #835 and #403 suites; earlier
    test modules overwrite slide-part registrations with mocks and do not
    restore them, which breaks ``Presentation(stream)`` reopen in the
    save/reopen round-trip assertions below.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
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
        CT.PML_NOTES_MASTER: NotesMasterPart,
        CT.PML_NOTES_SLIDE: NotesSlidePart,
        CT.PML_SLIDE: SlidePart,
        CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
        CT.PML_SLIDE_MASTER: SlideMasterPart,
        CT.DML_CHART: ChartPart,
        CT.JPEG: ImagePart,
        CT.PNG: ImagePart,
        CT.MP4: MediaPart,
        CT.SML_SHEET: EmbeddedXlsxPart,
    }
    PartFactory.part_type_for.update(expected)
    try:
        yield
    finally:
        PartFactory.part_type_for.clear()
        PartFactory.part_type_for.update(saved)


def _build_source_deck() -> PresentationClass:
    """Return a presentation with three chart slides, each with a distinct xlsx.

    Each chart is constructed from a unique ``(categories, values,
    series_name)`` triple, which means XlsxWriter emits a *different*
    workbook blob for every chart. That in turn lets the test distinguish
    cross-pollination (two charts resolving to the same workbook) from
    the correct behaviour (each chart keeps its own).
    """
    prs = Presentation()
    for categories, values, series_name in _CHART_DATA:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        data = CategoryChartData()
        data.categories = list(categories)
        data.add_series(series_name, values)
        slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(5),
            Inches(3),
            data,
        )
    return prs


def _first_chart_on(slide: Slide) -> Chart:
    for shape in slide.shapes:
        if getattr(shape, "has_chart", False):
            return shape.chart
    raise AssertionError("no chart on slide")


def _xlsx_part_of(slide: Slide) -> EmbeddedXlsxPart:
    """Return the :class:`EmbeddedXlsxPart` backing the chart on *slide*."""
    chart_part = cast(ChartPart, _first_chart_on(slide).part)
    xlsx_part = chart_part.chart_workbook.xlsx_part
    assert xlsx_part is not None, "chart has no embedded xlsx workbook"
    return xlsx_part


def _xlsx_blob_of(slide: Slide) -> bytes:
    """Return the raw bytes of the chart's embedded ``.xlsx`` workbook.

    This is what PowerPoint would surface from the chart's "Edit Data"
    dialog; comparing these blobs across charts is the sharpest test of
    "no cross-pollination".
    """
    return _xlsx_part_of(slide).blob


def _values_of(slide: Slide) -> tuple[float, ...]:
    chart = _first_chart_on(slide)
    return tuple(chart.plots[0].series[0].values)


def _save_and_reopen(prs: PresentationClass) -> PresentationClass:
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


class DescribeIssue742MergeXlsxChartsPreserveEachWorkbook(object):
    """#742 — close as resolved by #934 + #1036 (with #835 companion).

    Pins the "each merged chart keeps its own embedded ``.xlsx``" half of
    the cross-presentation chart-copy guarantee: after
    :meth:`Presentation.merge`, every chart in the target deck carries
    a distinct workbook blob, its series values match the source, and
    the invariant survives save+reopen.
    """

    # -- (1) every source chart enters the target with its own xlsx part ----

    def it_gives_each_merged_chart_a_distinct_embedded_xlsx_part(self, _restore_part_factory):
        """Chart -> EmbeddedXlsxPart relationships are 1:1 in the merged deck.

        Cross-references #934's F5 ``clone_embedded_xlsx`` guarantee: the
        merge must materialise a *fresh* :class:`EmbeddedXlsxPart` for
        every chart the cloner copies. If any two merged charts shared a
        part, PowerPoint's "Edit Data" on one would mutate the other.
        """
        source = _build_source_deck()

        target = Presentation()
        target.merge(source)

        xlsx_parts_by_chart = [_xlsx_part_of(sl) for sl in target.slides]
        # -- every chart's xlsx part is a distinct object --
        ids = [id(p) for p in xlsx_parts_by_chart]
        assert len(set(ids)) == len(ids), f"merged charts share EmbeddedXlsxPart instances: {ids}"

    def it_gives_each_merged_chart_a_distinct_xlsx_blob(self, _restore_part_factory):
        """No cross-pollination: each chart's workbook bytes are unique.

        Object identity (above) is necessary but not sufficient — two
        distinct parts could still carry identical bytes if the cloner
        mis-routed data. This asserts the bytes themselves differ, which
        is the strictest form of the "no cross-pollination" guarantee.
        """
        source = _build_source_deck()

        target = Presentation()
        target.merge(source)

        blobs = [_xlsx_blob_of(sl) for sl in target.slides]
        assert len(set(blobs)) == len(blobs), "merged charts collapsed onto the same xlsx blob"

    def it_preserves_each_source_chart_xlsx_blob_exactly(self, _restore_part_factory):
        """Each merged chart's xlsx blob matches its source slide's blob.

        The F5 cloner duplicates the raw xlsx bytes — the merged chart's
        embedded workbook should be byte-identical to the source chart's
        embedded workbook at the time of the merge.
        """
        source = _build_source_deck()
        source_blobs = [_xlsx_blob_of(sl) for sl in source.slides]

        target = Presentation()
        target.merge(source)

        target_blobs = [_xlsx_blob_of(sl) for sl in target.slides]
        assert target_blobs == source_blobs

    # -- (2) merged target does not reference any source-deck xlsx parts ----

    def it_does_not_share_xlsx_parts_with_the_source_package(self, _restore_part_factory):
        """Source and target packages own disjoint EmbeddedXlsxPart sets.

        Reaffirms the #835 cross-package-isolation guarantee from the
        #742 angle: after the merge, no :class:`EmbeddedXlsxPart` in the
        target package is the same instance as any in the source.
        """
        source = _build_source_deck()
        source_xlsx_ids = {
            id(p) for p in source.part.package.iter_parts() if isinstance(p, EmbeddedXlsxPart)
        }

        target = Presentation()
        target.merge(source)

        target_xlsx_ids = {
            id(p) for p in target.part.package.iter_parts() if isinstance(p, EmbeddedXlsxPart)
        }
        assert target_xlsx_ids
        assert not target_xlsx_ids.intersection(source_xlsx_ids)

    # -- (3) round-trip: save+reopen preserves per-chart xlsx identity -----

    def it_survives_save_and_reopen_with_distinct_xlsx_blobs(self, _restore_part_factory):
        """The per-chart-xlsx guarantee survives :meth:`Presentation.save`.

        After merge + save + reopen, every chart still has its own
        workbook, the blobs are still pairwise-distinct, and the series
        values PowerPoint would compute from each chart match the source
        deck's expectations.
        """
        source = _build_source_deck()
        expected_values = [_values_of(sl) for sl in source.slides]

        target = Presentation()
        target.merge(source)

        reopened = _save_and_reopen(target)

        # -- slide count preserved --
        assert len(reopened.slides) == len(_CHART_DATA)

        # -- every chart still has its own embedded xlsx part --
        xlsx_parts = [_xlsx_part_of(sl) for sl in reopened.slides]
        ids = [id(p) for p in xlsx_parts]
        assert len(set(ids)) == len(ids)

        # -- every chart's xlsx blob is still unique --
        blobs = [_xlsx_blob_of(sl) for sl in reopened.slides]
        assert len(set(blobs)) == len(blobs)

        # -- every chart's cached series values match the source --
        assert [_values_of(sl) for sl in reopened.slides] == expected_values
