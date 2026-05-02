# pyright: reportPrivateUsage=false

"""Regression test for issue #791 - delete a row from an existing table.

Issue #791 (https://github.com/scanny/python-pptx/issues/791) asked
for a way to delete a specific row from a table - in the reporter's
words, "I want to delete last 7th row from table." The reporter even
sketched the workaround ``table._tbl.remove(row._tr)`` that had been
circulating on issue #192.

That workaround bypassed every invariant python-pptx maintains
around a table (notably the graphic-frame height, which is derived
from the sum of authored row heights), so it could not be promoted
as-is. Issue #837 (Wave 3, ``feat/issue-837-table-row-delete``)
shipped the proper, supported counterpart:

  * :meth:`._Row.delete` detaches the row's ``a:tr`` from its parent
    ``a:tbl`` and recomputes the containing graphic-frame height.
  * :meth:`._RowCollection.remove` is the collection-level form,
    raising :class:`ValueError` if the caller passes a row from a
    different table.

#791 is therefore a duplicate of #837 and is verified-and-closed by
this suite. The scenarios below exercise the reporter's exact
workflow (a 7-row table, delete the 7th row, save, reopen, confirm
6 rows survive) plus the collection-level API and the error path a
naive "remove a foreign row" caller would trip into, pinning both
APIs against the kind of silent regression that would reopen the
issue.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.table import _Row, _RowCollection
from pptx.util import Emu, Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against pollution from tests that mutate `PartFactory.part_type_for`.

    See the identical fixture in ``tests/test_issue_937_table_cell_layout.py``
    for the rationale - other regression suites mock out the part registry
    and, if their teardown leaks, this suite's ``Presentation.save`` + reopen
    round-trips encounter a ``Mock`` in place of ``SlidePart`` and explode
    with ``AttributeError: Mock object has no attribute 'slide'``.
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
    }
    PartFactory.part_type_for.update(expected)
    try:
        yield
    finally:
        PartFactory.part_type_for.clear()
        PartFactory.part_type_for.update(saved)


def _build_seven_row_table(prs):
    """Return the graphic-frame shape for a 7x3 table mirroring the #791 scenario.

    Each cell is labelled ``r<row>c<col>`` so row identity is visible after
    mutation.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_table(
        rows=7,
        cols=3,
        left=Inches(1),
        top=Inches(1),
        width=Inches(6),
        height=Inches(3),
    )
    table = shape.table
    for r in range(7):
        for c in range(3):
            table.cell(r, c).text = f"r{r}c{c}"
    return shape


class DescribeIssue791DeleteTableRow(object):
    """#791 delete-row verify-and-close via #837's ``_Row.delete`` /
    ``_RowCollection.remove``.
    """

    # -- reporter's workflow -----------------------------------------------

    def it_deletes_the_seventh_row_the_reporter_asked_about(self, _restore_part_factory):
        # -- the #791 reporter's exact scenario: a 7-row table, delete the
        # -- 7th (last) row via the indexed-row API
        prs = Presentation()
        shape = _build_seven_row_table(prs)
        table = shape.table

        table.rows[6].delete()

        assert len(table.rows) == 6
        # -- the surviving rows are rows 0..5; row 6 is gone
        surviving_text = [[table.cell(r, c).text for c in range(3)] for r in range(6)]
        assert surviving_text == [[f"r{r}c{c}" for c in range(3)] for r in range(6)]

    def it_deletes_any_row_not_just_the_last(self, _restore_part_factory):
        # -- pin that row identity (cell text) moves correctly after an
        # -- interior delete - the reporter's snippet looked like it
        # -- would silently mis-aim if the authored row did not match
        prs = Presentation()
        shape = _build_seven_row_table(prs)
        table = shape.table

        # -- delete the 3rd row (index 2)
        table.rows[2].delete()

        assert len(table.rows) == 6
        assert [table.cell(r, 0).text for r in range(6)] == [
            "r0c0",
            "r1c0",
            "r3c0",
            "r4c0",
            "r5c0",
            "r6c0",
        ]

    # -- collection-level counterpart --------------------------------------

    def it_can_remove_a_row_via_rows_remove(self, _restore_part_factory):
        prs = Presentation()
        shape = _build_seven_row_table(prs)
        table = shape.table
        target = table.rows[3]

        table.rows.remove(target)

        assert len(table.rows) == 6
        assert [table.cell(r, 0).text for r in range(6)] == [
            "r0c0",
            "r1c0",
            "r2c0",
            "r4c0",
            "r5c0",
            "r6c0",
        ]

    def it_raises_when_removing_a_row_from_a_different_table(self, _restore_part_factory):
        # -- a naive caller who cached a row from another table must get
        # -- a clear ValueError instead of silently corrupting either
        # -- table
        prs = Presentation()
        shape_a = _build_seven_row_table(prs)
        shape_b = _build_seven_row_table(prs)
        foreign_row = shape_b.table.rows[0]

        with pytest.raises(ValueError, match="row is not a member of this table"):
            shape_a.table.rows.remove(foreign_row)
        # -- both tables still have all their rows
        assert len(shape_a.table.rows) == 7
        assert len(shape_b.table.rows) == 7

    # -- graphic-frame height invariant (the workaround couldn't maintain) --

    def it_shrinks_the_graphic_frame_height_by_the_deleted_row_height(self, _restore_part_factory):
        # -- the #791 one-liner workaround left the graphic-frame cy
        # -- stale; the supported API recomputes it as sum(row heights).
        prs = Presentation()
        shape = _build_seven_row_table(prs)
        table = shape.table
        height_before = Emu(shape.height)
        deleted_row_height = Emu(table.rows[6].height)

        table.rows[6].delete()

        assert Emu(shape.height) == height_before - deleted_row_height
        assert Emu(shape.height) == Emu(sum(row.height for row in table.rows))

    def it_also_shrinks_the_frame_when_removed_via_collection(self, _restore_part_factory):
        prs = Presentation()
        shape = _build_seven_row_table(prs)
        table = shape.table
        height_before = Emu(shape.height)
        target = table.rows[0]
        deleted_row_height = Emu(target.height)

        table.rows.remove(target)

        assert Emu(shape.height) == height_before - deleted_row_height

    # -- XML-shape invariant -----------------------------------------------

    def it_removes_the_a_tr_from_the_a_tbl(self, _restore_part_factory):
        prs = Presentation()
        shape = _build_seven_row_table(prs)
        table = shape.table
        tbl = table._tbl
        tr_to_delete = table.rows[6]._tr

        table.rows[6].delete()

        assert tr_to_delete not in tbl.tr_lst
        assert len(tbl.tr_lst) == 6

    # -- round-trip --------------------------------------------------------

    def it_round_trips_the_deleted_row_through_save_reopen(self, _restore_part_factory):
        # -- the reporter's end goal: save a file with one row fewer and
        # -- have PowerPoint (modelled here by Presentation(buf)) see the
        # -- same shape.
        prs = Presentation()
        shape = _build_seven_row_table(prs)
        table = shape.table
        table.rows[6].delete()
        expected_height = shape.height

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        shape2 = next(s for s in prs2.slides[0].shapes if s.has_table)
        table2 = shape2.table

        assert len(table2.rows) == 6
        assert [table2.cell(r, 0).text for r in range(6)] == [f"r{r}c0" for r in range(6)]
        # -- the authored frame height survives the round-trip
        assert shape2.height == expected_height

    def it_round_trips_multiple_deletes_in_sequence(self, _restore_part_factory):
        # -- #791's two snippets were both "remove one row" helpers; a
        # -- real caller typically deletes more than one. Pin that a
        # -- sequence of deletes + reopen yields exactly the rows that
        # -- were not deleted.
        prs = Presentation()
        shape = _build_seven_row_table(prs)
        table = shape.table

        # -- delete rows 6, 4, and 0 (from tail to head so earlier indices
        # -- keep pointing at the intended row)
        for idx in (6, 4, 0):
            table.rows[idx].delete()

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        table2 = next(s for s in prs2.slides[0].shapes if s.has_table).table

        assert len(table2.rows) == 4
        assert [table2.cell(r, 0).text for r in range(4)] == [
            "r1c0",
            "r2c0",
            "r3c0",
            "r5c0",
        ]

    # -- public API surface ------------------------------------------------

    def it_exposes_delete_on_Row(self):
        # -- belt-and-braces: pin that the attribute the reporter needed
        # -- actually exists on the class, not just on instances under
        # -- some fixture.
        assert hasattr(_Row, "delete")
        assert callable(_Row.delete)

    def it_exposes_remove_on_RowCollection(self):
        assert hasattr(_RowCollection, "remove")
        assert callable(_RowCollection.remove)
