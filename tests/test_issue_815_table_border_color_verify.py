"""Regression test for issue #815 — change border color of a table.

Issue #815 (https://github.com/scanny/python-pptx/issues/815) asked how to
change the border color of a table. In PowerPoint, table borders are
per-cell constructs (``a:lnL`` / ``a:lnR`` / ``a:lnT`` / ``a:lnB`` on each
``a:tc/a:tcPr``, plus the two diagonals ``a:lnTlToBr`` / ``a:lnBlToTr``) —
there is no table-level "border color" element. The fork closed the gap
with feature issue #71 (``feat/issue-71-cell-borders``, Wave 3), which
added six ``LineFormat`` properties on :class:`pptx.table._Cell`:

    * ``_Cell.border_left``           → ``a:lnL``
    * ``_Cell.border_right``          → ``a:lnR``
    * ``_Cell.border_top``            → ``a:lnT``
    * ``_Cell.border_bottom``         → ``a:lnB``
    * ``_Cell.border_diagonal_down``  → ``a:lnTlToBr``
    * ``_Cell.border_diagonal_up``    → ``a:lnBlToTr``

Each property returns a |LineFormat| supporting ``color.rgb``, ``width``,
and ``dash_style``, matching the rest of the DrawingML line surface. Paired
with :meth:`Table.iter_cells`, the bulk-idiom

.. code-block:: python

    for cell in table.iter_cells():
        for border in (cell.border_left, cell.border_right,
                       cell.border_top, cell.border_bottom):
            border.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            border.width = Pt(0.75)

applies a uniform border colour to every cell in the table — which is the
answer to the #815 reporter's question.

The scenarios below pin (a) per-side color/width assignment on a single
cell, (b) diagonal borders, (c) the table-wide bulk idiom via
``iter_cells``, and (d) the full save + reopen round-trip so the XML
persists — guarding against silent regression that would reopen the issue.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt


def _fresh_table(rows=3, cols=3):
    """Return ``(prs, table)`` for a freshly-authored table on a blank slide.

    Uses slide layout 6 (the blank layout) so there are no placeholder
    shapes to filter past when reopening in the round-trip tests.
    """
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_table(
        rows=rows,
        cols=cols,
        left=Inches(1),
        top=Inches(1),
        width=Inches(6),
        height=Inches(2),
    )
    return prs, shape.table


class DescribeIssue815TableBorderColor(object):
    """End-to-end regression coverage for issue #815.

    The capability was shipped under feature #71 (``feat/issue-71-cell-borders``,
    Wave 3). These scenarios pin the reporter's exact workflow — "how do I
    change the border color of my table?" — through authoring, bulk
    application, diagonals, and save + reopen.
    """

    # -- public-surface smoke: the six ``_Cell.border_*`` LineFormat properties
    # -- must exist, because that is the whole answer to #815 ---------------

    def it_exposes_four_edge_and_two_diagonal_border_properties_on_Cell(self):
        from pptx.dml.line import LineFormat
        from pptx.table import _Cell

        _, table = _fresh_table(rows=1, cols=1)
        cell = table.cell(0, 0)

        # -- all six properties are live attributes of the runtime class
        for name in (
            "border_left",
            "border_right",
            "border_top",
            "border_bottom",
            "border_diagonal_down",
            "border_diagonal_up",
        ):
            assert hasattr(_Cell, name), f"_Cell is missing {name!r}"
            assert isinstance(getattr(cell, name), LineFormat)

    # -- the reporter's question: set per-side border color ---------------

    def it_writes_border_left_color_to_the_cell_tcPr(self):
        _, table = _fresh_table(rows=1, cols=1)
        cell = table.cell(0, 0)

        cell.border_left.color.rgb = RGBColor(0xC0, 0x00, 0x00)

        # -- `a:lnL` is now present under `a:tcPr/a:lnL/a:solidFill/a:srgbClr`
        tcPr = cell._tc.find(qn("a:tcPr"))
        assert tcPr is not None
        lnL = tcPr.find(qn("a:lnL"))
        assert lnL is not None
        srgbClr = lnL.find(f"{qn('a:solidFill')}/{qn('a:srgbClr')}")
        assert srgbClr is not None
        assert srgbClr.get("val") == "C00000"

    def it_writes_all_four_edge_border_colors_on_a_single_cell(self):
        _, table = _fresh_table(rows=1, cols=1)
        cell = table.cell(0, 0)
        color = RGBColor(0x44, 0x44, 0x44)

        for border in (
            cell.border_left,
            cell.border_right,
            cell.border_top,
            cell.border_bottom,
        ):
            border.color.rgb = color
            border.width = Pt(1.5)

        tcPr = cell._tc.find(qn("a:tcPr"))
        assert tcPr is not None
        # -- each per-side `a:ln?` element is present with the same sRGB color
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            ln = tcPr.find(qn(tag))
            assert ln is not None, f"{tag} not written"
            srgb = ln.find(f"{qn('a:solidFill')}/{qn('a:srgbClr')}")
            assert srgb is not None
            assert srgb.get("val") == "444444"
            # -- 1.5 pt = 19050 EMU (12700 EMU/pt)
            assert int(ln.get("w")) == 19050

    def it_writes_both_diagonal_border_colors_on_a_single_cell(self):
        _, table = _fresh_table(rows=1, cols=1)
        cell = table.cell(0, 0)

        cell.border_diagonal_down.color.rgb = RGBColor(0x00, 0x66, 0xCC)
        cell.border_diagonal_up.color.rgb = RGBColor(0xCC, 0x00, 0x66)

        tcPr = cell._tc.find(qn("a:tcPr"))
        assert tcPr is not None
        # -- top-left→bottom-right diagonal
        lnTlToBr = tcPr.find(qn("a:lnTlToBr"))
        assert lnTlToBr is not None
        tl_srgb = lnTlToBr.find(f"{qn('a:solidFill')}/{qn('a:srgbClr')}")
        assert tl_srgb is not None
        assert tl_srgb.get("val") == "0066CC"
        # -- bottom-left→top-right diagonal
        lnBlToTr = tcPr.find(qn("a:lnBlToTr"))
        assert lnBlToTr is not None
        bl_srgb = lnBlToTr.find(f"{qn('a:solidFill')}/{qn('a:srgbClr')}")
        assert bl_srgb is not None
        assert bl_srgb.get("val") == "CC0066"

    def it_accepts_a_dash_style_on_a_border(self):
        _, table = _fresh_table(rows=1, cols=1)
        cell = table.cell(0, 0)

        cell.border_bottom.color.rgb = RGBColor(0x80, 0x80, 0x80)
        cell.border_bottom.width = Pt(1)
        cell.border_bottom.dash_style = MSO_LINE.DASH

        assert cell.border_bottom.dash_style == MSO_LINE.DASH
        # -- width round-trips as an EMU-typed Length
        assert Emu(cell.border_bottom.width) == Pt(1)

    # -- the bulk idiom: color every border in the table ------------------

    def it_applies_a_uniform_border_color_to_every_cell_via_iter_cells(self):
        # -- the exact recipe cited in docs/user/table.rst: one loop over
        # -- `table.iter_cells()`, four border assignments per cell.
        _, table = _fresh_table(rows=3, cols=4)
        color = RGBColor(0x95, 0xB3, 0xD7)

        for cell in table.iter_cells():
            for border in (
                cell.border_left,
                cell.border_right,
                cell.border_top,
                cell.border_bottom,
            ):
                border.color.rgb = color
                border.width = Pt(0.75)

        # -- every one of 12 cells carries all four explicit borders with
        # -- the uniform colour and width
        for cell in table.iter_cells():
            for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                tcPr = cell._tc.find(qn("a:tcPr"))
                assert tcPr is not None
                ln = tcPr.find(qn(tag))
                assert ln is not None, f"{tag} missing on cell ({cell.row_idx},{cell.col_idx})"
                srgb = ln.find(f"{qn('a:solidFill')}/{qn('a:srgbClr')}")
                assert srgb is not None
                assert srgb.get("val") == "95B3D7"

    # -- round-trip: save + reopen preserves every border ------------------

    def it_round_trips_per_cell_border_color_through_save_reopen(self):
        prs, table = _fresh_table(rows=2, cols=2)
        color = RGBColor(0xC0, 0x00, 0x00)
        diag = RGBColor(0x00, 0x80, 0x00)

        # -- every edge on every cell gets the same red border
        for cell in table.iter_cells():
            for border in (
                cell.border_left,
                cell.border_right,
                cell.border_top,
                cell.border_bottom,
            ):
                border.color.rgb = color
                border.width = Pt(1)
        # -- the (0, 0) cell additionally gets a green top-left→bottom-right diagonal
        table.cell(0, 0).border_diagonal_down.color.rgb = diag
        table.cell(0, 0).border_diagonal_down.width = Pt(0.75)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        shape2 = next(s for s in prs2.slides[0].shapes if s.has_table)
        table2 = shape2.table

        # -- every edge on every cell reports the authored red color
        for cell in table2.iter_cells():
            for side in ("border_left", "border_right", "border_top", "border_bottom"):
                line = getattr(cell, side)
                assert (
                    line.color.rgb == color
                ), f"cell ({cell.row_idx},{cell.col_idx}).{side} did not round-trip"
                # -- width survives the round-trip (Emu-typed Length)
                assert Emu(line.width) == Pt(1)
        # -- the diagonal persists on (0, 0)
        assert table2.cell(0, 0).border_diagonal_down.color.rgb == diag
        assert Emu(table2.cell(0, 0).border_diagonal_down.width) == Pt(0.75)

    def it_round_trips_a_mixed_edge_and_diagonal_styling(self):
        # -- tighter-scoped round-trip pinned for the diagonal pair, which
        # -- the OOXML spec lists as "frequently emitted incorrectly" and
        # -- therefore the most likely to silently drop on reopen.
        prs, table = _fresh_table(rows=1, cols=1)
        cell = table.cell(0, 0)
        cell.border_diagonal_down.color.rgb = RGBColor(0x00, 0x66, 0xCC)
        cell.border_diagonal_down.width = Pt(2)
        cell.border_diagonal_up.color.rgb = RGBColor(0xCC, 0x00, 0x66)
        cell.border_diagonal_up.width = Pt(2)
        cell.border_top.color.rgb = RGBColor(0xFF, 0xFF, 0x00)
        cell.border_top.width = Pt(1)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        shape2 = next(s for s in prs2.slides[0].shapes if s.has_table)
        cell2 = shape2.table.cell(0, 0)

        assert cell2.border_diagonal_down.color.rgb == RGBColor(0x00, 0x66, 0xCC)
        assert cell2.border_diagonal_up.color.rgb == RGBColor(0xCC, 0x00, 0x66)
        assert cell2.border_top.color.rgb == RGBColor(0xFF, 0xFF, 0x00)
        assert Emu(cell2.border_diagonal_down.width) == Pt(2)
        assert Emu(cell2.border_diagonal_up.width) == Pt(2)
        assert Emu(cell2.border_top.width) == Pt(1)
