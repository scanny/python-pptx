# pyright: reportPrivateUsage=false

"""Regression test for issue #403 — feature: merge presentation file.

Issue #403 (https://github.com/scanny/python-pptx/issues/403) asks for a
``merge`` function that combines two filled ``.pptx`` files into a single
output file. The reporter's workflow is:

    "I could make a presentation template manually, save it as **a.pptx**,
    make another presentation template b, save it as **b.pptx**. So that I
    could filling some stuffs into **a.pptx** and **b.pptx** by coding then
    a merge function that combines the two filled file into a single
    **c.pptx** is what I really want."

Wave 7 #934 shipped :meth:`Presentation.merge` — appends every slide of
``other_presentation`` to the receiver, full-fidelity: image, media, chart
(with a *distinct* embedded workbook), OLE object, embedded-package, and
external-hyperlink parts are all materialised in the target package via
the Foundation F1 (``PartRelationshipCloner``) + F5
(``clone_embedded_xlsx``) cloning pipeline that #934 promoted to the
default for :meth:`Slides.add_slide_from_external`. The exact combine-two-
filled-files-into-one workflow the #403 reporter asked for is now a three-
line recipe::

    a = Presentation("a.pptx")
    a.merge(Presentation("b.pptx"))
    a.save("c.pptx")

This regression suite pins the end-to-end behaviour for the #403 scenario
so the issue can be closed as *resolved by #934*:

  * Two filled presentations can be combined into a single output file via
    :meth:`Presentation.merge` + :meth:`Presentation.save`.
  * Slide content (text, pictures, charts, tables) round-trips through the
    merge + save + reopen cycle.
  * Charts materialise their own :class:`EmbeddedXlsxPart` in the merged
    package so PowerPoint's "Edit Data" dialog keeps working on both the
    original and the merged copy.
  * The source presentation is *not* mutated by the merge (the reporter's
    "fill two templates separately" workflow requires each input file to
    survive the merge intact).
  * Merging three or more decks works by chaining :meth:`Presentation.merge`
    calls — the #403 thread's implicit "merge N files" expectation.
"""

from __future__ import annotations

import io
from typing import TYPE_CHECKING

import pytest

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

if TYPE_CHECKING:
    from pptx.presentation import Presentation as PresentationClass
    from pptx.slide import Slide


@pytest.fixture
def _restore_part_factory():
    """Guard against ``PartFactory.part_type_for`` pollution from other tests.

    Mirrors the identical guard used by ``tests/test_issue_175_*.py`` and
    siblings; earlier test modules overwrite slide-part registrations with
    mocks and do not restore them, which breaks ``Presentation(stream)``
    reopen in save/reopen round-trip assertions.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.parts.embeddedpackage import EmbeddedXlsxPart
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


def _absjoin(*args: str) -> str:
    from .unitutil.file import absjoin, test_file_dir

    return absjoin(test_file_dir, *args)


def _fill_deck_a(prs: PresentationClass) -> None:
    """Populate `prs` with two picture slides mimicking the reporter's "a.pptx" template.

    Both slides carry a picture shape — the simplest "filled template"
    content that exercises :class:`ImagePart` materialisation across the
    merge, matching the "filling some stuffs into a.pptx" phase of the
    #403 workflow.
    """
    for _ in range(2):
        sl = prs.slides.add_slide(prs.slide_layouts[5])
        sl.shapes.add_picture(_absjoin("python-icon.jpeg"), Inches(1), Inches(1), height=Inches(1))


def _fill_deck_b(prs: PresentationClass) -> None:
    """Populate `prs` with two slides mimicking the reporter's "b.pptx" template.

    Slide 1 carries a chart; slide 2 carries a table. Between them they
    exercise the #934 cross-package code paths the #403 feature depends on:
    ``ChartPart.clone_from`` (distinct embedded workbook), and ordinary
    shape-tree cloning for the table XML.
    """
    chart_slide = prs.slides.add_slide(prs.slide_layouts[5])
    data = CategoryChartData()
    data.categories = ["Q1", "Q2", "Q3"]
    data.add_series("Series 1", (10.0, 20.0, 30.0))
    chart_slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(5),
        Inches(3),
        data,
    )

    table_slide = prs.slides.add_slide(prs.slide_layouts[5])
    table_slide.shapes.add_table(
        rows=2,
        cols=2,
        left=Inches(1),
        top=Inches(1),
        width=Inches(4),
        height=Inches(1),
    )


def _has_picture(slide: Slide) -> bool:
    return any(sh.shape_type == MSO_SHAPE_TYPE.PICTURE for sh in slide.shapes)


def _has_chart(slide: Slide) -> bool:
    return any(getattr(sh, "has_chart", False) for sh in slide.shapes)


def _has_table(slide: Slide) -> bool:
    return any(getattr(sh, "has_table", False) for sh in slide.shapes)


class DescribeIssue403MergePresentationFile(object):
    """#403 — close as resolved by #934.

    Pins the "combine two filled .pptx files into a single .pptx" workflow
    that the #403 reporter asked for, including save/reopen round-trip,
    multi-deck chaining, and the distinct-embedded-workbook guarantee for
    chart-bearing source slides.
    """

    # -- (1) the headline #403 scenario: combine a.pptx + b.pptx into c.pptx --

    def it_combines_two_filled_files_into_a_single_output(self, _restore_part_factory):
        """The exact workflow from the #403 issue body.

        Fill two separate presentations, serialise both to bytes (simulating
        ``a.pptx``/``b.pptx`` on disk), reopen the first as the target, merge
        in the second, save as ``c.pptx``, and verify the combined file
        carries every slide from both inputs.
        """
        # -- (a) build and serialise the two filled templates --
        prs_a = Presentation()
        _fill_deck_a(prs_a)
        a_bytes = io.BytesIO()
        prs_a.save(a_bytes)

        prs_b = Presentation()
        _fill_deck_b(prs_b)
        b_bytes = io.BytesIO()
        prs_b.save(b_bytes)

        # -- (b) reopen both and perform the merge, exactly as the reporter
        # -- would from the filesystem: `c = a; c.merge(b); c.save("c.pptx")`.
        a_bytes.seek(0)
        b_bytes.seek(0)
        target = Presentation(a_bytes)
        other = Presentation(b_bytes)
        a_slide_count = len(target.slides)
        b_slide_count = len(other.slides)

        appended = target.merge(other)

        # -- (c) assertions: combined slide count + appended return value --
        assert len(appended) == b_slide_count
        assert len(target.slides) == a_slide_count + b_slide_count

        # -- (d) save the combined deck as "c.pptx" and verify it reopens --
        c_bytes = io.BytesIO()
        target.save(c_bytes)
        c_bytes.seek(0)
        reopened = Presentation(c_bytes)

        assert len(reopened.slides) == a_slide_count + b_slide_count

        # -- (e) a.pptx content survives in the combined file --
        assert _has_picture(reopened.slides[1])  # a's picture slide

        # -- (f) b.pptx content survives in the combined file --
        assert _has_chart(reopened.slides[a_slide_count])  # b's chart slide
        assert _has_table(reopened.slides[a_slide_count + 1])  # b's table slide

    # -- (2) source file is not mutated by the merge ------------------------

    def it_does_not_mutate_the_source_presentation(self, _restore_part_factory):
        """The #403 workflow fills each template separately, so the source
        file must be untouched after a merge — otherwise chaining merges
        (or re-using ``b.pptx`` for another combine) would double-count.
        """
        target = Presentation()
        _fill_deck_a(target)

        source = Presentation()
        _fill_deck_b(source)
        source_slide_count_before = len(source.slides)
        source_shape_counts_before = [len(sl.shapes) for sl in source.slides]

        target.merge(source)

        assert len(source.slides) == source_slide_count_before
        assert [len(sl.shapes) for sl in source.slides] == source_shape_counts_before

    # -- (3) chart slides get a *distinct* embedded workbook in the merged deck

    def it_gives_merged_chart_slides_their_own_embedded_workbook(self, _restore_part_factory):
        """#934's headline full-fidelity guarantee for the #403 use case.

        A chart slide copied via the merge pipeline must end up with a
        :class:`EmbeddedXlsxPart` in the target package that is *not* one
        of the source's xlsx parts — if the target shared the source's
        workbook part, PowerPoint's "Edit Data" dialog would mutate both
        decks at once. Uses ``id()``-based identity comparison (matching the
        Wave 7 #934 unit-test pattern) rather than counts, which are noisy
        in the pre-save in-memory package where intermediate xlsx parts
        from the chart-creation pipeline are still reachable.
        """
        from pptx.parts.embeddedpackage import EmbeddedXlsxPart

        source = Presentation()
        _fill_deck_b(source)  # one chart slide + one table slide
        source_xlsx_ids = {
            id(p) for p in source.part.package.iter_parts() if isinstance(p, EmbeddedXlsxPart)
        }

        target = Presentation()
        target.merge(source)

        target_xlsx_parts = [
            p for p in target.part.package.iter_parts() if isinstance(p, EmbeddedXlsxPart)
        ]
        # -- the merge materialised at least one xlsx in the target package --
        assert target_xlsx_parts
        # -- cross-package isolation: no target xlsx part is a source xlsx --
        target_xlsx_ids = {id(p) for p in target_xlsx_parts}
        assert not target_xlsx_ids.intersection(source_xlsx_ids)

    # -- (4) chaining: merge three or more decks into one output ------------

    def it_chains_merges_to_combine_three_or_more_decks(self, _restore_part_factory):
        """The #403 thread's implicit "merge N files" expectation.

        ``merge`` is idempotent-over-chaining: calling it repeatedly appends
        each source deck's slides to the receiver in order, matching the
        PowerPoint "reuse slides" gesture for multiple source files.
        """
        prs_a = Presentation()
        _fill_deck_a(prs_a)  # 2 slides
        prs_b = Presentation()
        _fill_deck_b(prs_b)  # 2 slides
        prs_c = Presentation()
        _fill_deck_a(prs_c)  # 2 more slides — reuses the deck-A shape --

        target = Presentation()
        target_pre = len(target.slides)

        target.merge(prs_a)
        target.merge(prs_b)
        target.merge(prs_c)

        assert len(target.slides) == target_pre + 6

        # -- round-trip the combined deck to confirm the three-way merge
        # -- survives save + reopen with every slide intact --
        buf = io.BytesIO()
        target.save(buf)
        buf.seek(0)
        reopened = Presentation(buf)
        assert len(reopened.slides) == target_pre + 6

    # -- (5) sanity: reject self-merge and non-presentation arguments -------

    def it_rejects_merging_a_presentation_into_itself(self, _restore_part_factory):
        prs = Presentation()
        _fill_deck_a(prs)
        with pytest.raises(ValueError, match="itself"):
            prs.merge(prs)

    def it_rejects_merging_a_non_presentation_argument(self, _restore_part_factory):
        prs = Presentation()
        with pytest.raises(TypeError, match="must be a Presentation"):
            prs.merge("b.pptx")  # type: ignore[arg-type]
