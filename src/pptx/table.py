"""Table-related objects such as Table and Cell."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from pptx.dml.fill import FillFormat
from pptx.dml.line import LineFormat
from pptx.oxml.table import TcRange
from pptx.shapes import Subshape
from pptx.text.text import TextFrame
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.enum.text import MSO_VERTICAL_ANCHOR
    from pptx.oxml.shapes.shared import CT_LineProperties
    from pptx.oxml.table import CT_Table, CT_TableCell, CT_TableCol, CT_TableRow
    from pptx.parts.slide import BaseSlidePart
    from pptx.shapes.graphfrm import GraphicFrame
    from pptx.types import ProvidesPart
    from pptx.util import Length


class Table(object):
    """A DrawingML table object.

    Not intended to be constructed directly, use
    :meth:`.Slide.shapes.add_table` to add a table to a slide.
    """

    def __init__(self, tbl: CT_Table, graphic_frame: GraphicFrame):
        super(Table, self).__init__()
        self._tbl = tbl
        self._graphic_frame = graphic_frame

    def cell(self, row_idx: int, col_idx: int) -> _Cell:
        """Return cell at `row_idx`, `col_idx`.

        Return value is an instance of |_Cell|. `row_idx` and `col_idx` are zero-based, e.g.
        cell(0, 0) is the top, left cell in the table.
        """
        return _Cell(self._tbl.tc(row_idx, col_idx), self)

    @lazyproperty
    def columns(self) -> _ColumnCollection:
        """|_ColumnCollection| instance for this table.

        Provides access to |_Column| objects representing the table's columns. |_Column| objects
        are accessed using list notation, e.g. `col = tbl.columns[0]`.
        """
        return _ColumnCollection(self._tbl, self)

    @property
    def first_col(self) -> bool:
        """When `True`, indicates first column should have distinct formatting.

        Read/write. Distinct formatting is used, for example, when the first column contains row
        headings (is a side-heading column).
        """
        return self._tbl.firstCol

    @first_col.setter
    def first_col(self, value: bool):
        self._tbl.firstCol = value

    @property
    def first_row(self) -> bool:
        """When `True`, indicates first row should have distinct formatting.

        Read/write. Distinct formatting is used, for example, when the first row contains column
        headings.
        """
        return self._tbl.firstRow

    @first_row.setter
    def first_row(self, value: bool):
        self._tbl.firstRow = value

    @property
    def horz_banding(self) -> bool:
        """When `True`, indicates rows should have alternating shading.

        Read/write. Used to allow rows to be traversed more easily without losing track of which
        row is being read.
        """
        return self._tbl.bandRow

    @horz_banding.setter
    def horz_banding(self, value: bool):
        self._tbl.bandRow = value

    def iter_cells(self) -> Iterator[_Cell]:
        """Generate _Cell object for each cell in this table.

        Each grid cell is generated in left-to-right, top-to-bottom order.
        """
        return (_Cell(tc, self) for tc in self._tbl.iter_tcs())

    @property
    def last_col(self) -> bool:
        """When `True`, indicates the rightmost column should have distinct formatting.

        Read/write. Used, for example, when a row totals column appears at the far right of the
        table.
        """
        return self._tbl.lastCol

    @last_col.setter
    def last_col(self, value: bool):
        self._tbl.lastCol = value

    @property
    def last_row(self) -> bool:
        """When `True`, indicates the bottom row should have distinct formatting.

        Read/write. Used, for example, when a totals row appears as the bottom row.
        """
        return self._tbl.lastRow

    @last_row.setter
    def last_row(self, value: bool):
        self._tbl.lastRow = value

    def notify_height_changed(self) -> None:
        """Called by a row when its height changes.

        Triggers the graphic frame to recalculate its total height (as the sum of the row
        heights).

        .. note::
           The resulting graphic-frame height is the sum of the authored row heights, which
           PowerPoint treats as a *minimum* for each row. PowerPoint will grow a row as needed
           to fit its text content when it opens and lays out the slide, but `python-pptx`
           cannot perform that layout calculation. Consequently, the height reported by
           :attr:`.GraphicFrame.height` (and computed here) may be less than the rendered
           table height until the file is opened and saved by PowerPoint. See issue #296 and
           the "Table height and row height" section of the user guide.
        """
        new_table_height = Emu(sum([row.height for row in self.rows]))
        self._graphic_frame.height = new_table_height

    def notify_width_changed(self) -> None:
        """Called by a column when its width changes.

        Triggers the graphic frame to recalculate its total width (as the sum of the column
        widths).
        """
        new_table_width = Emu(sum([col.width for col in self.columns]))
        self._graphic_frame.width = new_table_width

    @property
    def part(self) -> BaseSlidePart:
        """The package part containing this table."""
        return self._graphic_frame.part

    @lazyproperty
    def rows(self):
        """|_RowCollection| instance for this table.

        Provides access to |_Row| objects representing the table's rows. |_Row| objects are
        accessed using list notation, e.g. `col = tbl.rows[0]`.
        """
        return _RowCollection(self._tbl, self)

    @property
    def style_id(self) -> str | None:
        """GUID of the table style referenced by this table, or `None`.

        The value is a string like ``"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"``. PowerPoint
        ships with a fixed set of built-in table styles and stores the selection as a GUID
        reference in the `a:tblPr/a:tableStyleId` element of the table XML. This property
        is read/write; assigning `None` removes the style reference, causing the host
        application to fall back to its default table style.

        Note the referenced style must exist in the presentation's `tableStyles.xml` part
        (or be one of PowerPoint's built-in GUIDs) for the assignment to take effect in
        the rendered output. python-pptx does not validate the GUID against that part and
        does not currently expose a typed, by-name API for selecting built-in styles; that
        is deferred to a future release (see issue #27).
        """
        return self._tbl.tableStyleId

    @style_id.setter
    def style_id(self, value: str | None) -> None:
        self._tbl.tableStyleId = value

    @property
    def vert_banding(self) -> bool:
        """When `True`, indicates columns should have alternating shading.

        Read/write. Used to allow columns to be traversed more easily without losing track of
        which column is being read.
        """
        return self._tbl.bandCol

    @vert_banding.setter
    def vert_banding(self, value: bool):
        self._tbl.bandCol = value


class _Cell(Subshape):
    """Table cell"""

    def __init__(self, tc: CT_TableCell, parent: ProvidesPart):
        super(_Cell, self).__init__(parent)
        self._tc = tc

    def __eq__(self, other: object) -> bool:
        """|True| if this object proxies the same element as `other`.

        Equality for proxy objects is defined as referring to the same XML element, whether or not
        they are the same proxy object instance.
        """
        if not isinstance(other, type(self)):
            return False
        return self._tc is other._tc

    def __ne__(self, other: object) -> bool:
        if not isinstance(other, type(self)):
            return True
        return self._tc is not other._tc

    @lazyproperty
    def border_bottom(self) -> LineFormat:
        """|LineFormat| for the bottom border of this cell (`a:lnB`).

        Assigning a color, width, or dash-style to this line causes the cell
        to draw an explicit bottom border in that style, overriding any
        border inherited from the table style. Reading properties from the
        returned |LineFormat| reports the explicitly-applied values; values
        inherited from a table style are not reported here.
        """
        return LineFormat(_CellBorder(self._tc, "lnB"))

    @lazyproperty
    def border_diagonal_down(self) -> LineFormat:
        """|LineFormat| for the top-left to bottom-right diagonal (`a:lnTlToBr`).

        Diagonal borders are drawn across the interior of the cell rather
        than along one of its edges; PowerPoint exposes them in the table
        styles dialog as "Borders > Diagonal Down". Behaves otherwise the
        same as the edge-border properties.
        """
        return LineFormat(_CellBorder(self._tc, "lnTlToBr"))

    @lazyproperty
    def border_diagonal_up(self) -> LineFormat:
        """|LineFormat| for the bottom-left to top-right diagonal (`a:lnBlToTr`).

        Diagonal borders are drawn across the interior of the cell rather
        than along one of its edges; PowerPoint exposes them in the table
        styles dialog as "Borders > Diagonal Up".
        """
        return LineFormat(_CellBorder(self._tc, "lnBlToTr"))

    @lazyproperty
    def border_left(self) -> LineFormat:
        """|LineFormat| for the left border of this cell (`a:lnL`).

        Assigning a color, width, or dash-style to this line causes the cell
        to draw an explicit left border in that style, overriding any border
        inherited from the table style. Reading properties from the returned
        |LineFormat| reports the explicitly-applied values; values inherited
        from a table style are not reported here.
        """
        return LineFormat(_CellBorder(self._tc, "lnL"))

    @lazyproperty
    def border_right(self) -> LineFormat:
        """|LineFormat| for the right border of this cell (`a:lnR`).

        See :attr:`border_left` for semantics.
        """
        return LineFormat(_CellBorder(self._tc, "lnR"))

    @lazyproperty
    def border_top(self) -> LineFormat:
        """|LineFormat| for the top border of this cell (`a:lnT`).

        See :attr:`border_left` for semantics.
        """
        return LineFormat(_CellBorder(self._tc, "lnT"))

    @lazyproperty
    def fill(self) -> FillFormat:
        """|FillFormat| instance for this cell.

        Provides access to fill properties such as foreground color.
        """
        tcPr = self._tc.get_or_add_tcPr()
        return FillFormat.from_fill_parent(tcPr)

    @property
    def col_idx(self) -> int:
        """Zero-based column index of this cell in its table.

        Reflects the position of this cell's `a:tc` element among its siblings in the containing
        `a:tr` row element. Read-only.
        """
        return self._tc.col_idx

    @property
    def is_merge_origin(self) -> bool:
        """True if this cell is the top-left grid cell in a merged cell."""
        return self._tc.is_merge_origin

    @property
    def is_spanned(self) -> bool:
        """True if this cell is spanned by a merge-origin cell.

        A merge-origin cell "spans" the other grid cells in its merge range, consuming their area
        and "shadowing" the spanned grid cells.

        Note this value is |False| for a merge-origin cell. A merge-origin cell spans other grid
        cells, but is not itself a spanned cell.
        """
        return self._tc.is_spanned

    @property
    def margin_left(self) -> Length:
        """Left margin of cells.

        Read/write. If assigned |None|, the default value is used, 0.1 inches for left and right
        margins and 0.05 inches for top and bottom.
        """
        return self._tc.marL

    @margin_left.setter
    def margin_left(self, margin_left: Length | None):
        self._validate_margin_value(margin_left)
        self._tc.marL = margin_left

    @property
    def margin_right(self) -> Length:
        """Right margin of cell."""
        return self._tc.marR

    @margin_right.setter
    def margin_right(self, margin_right: Length | None):
        self._validate_margin_value(margin_right)
        self._tc.marR = margin_right

    @property
    def margin_top(self) -> Length:
        """Top margin of cell."""
        return self._tc.marT

    @margin_top.setter
    def margin_top(self, margin_top: Length | None):
        self._validate_margin_value(margin_top)
        self._tc.marT = margin_top

    @property
    def margin_bottom(self) -> Length:
        """Bottom margin of cell."""
        return self._tc.marB

    @margin_bottom.setter
    def margin_bottom(self, margin_bottom: Length | None):
        self._validate_margin_value(margin_bottom)
        self._tc.marB = margin_bottom

    def merge(self, other_cell: _Cell) -> None:
        """Create merged cell from this cell to `other_cell`.

        This cell and `other_cell` specify opposite corners of the merged cell range. Either
        diagonal of the cell region may be specified in either order, e.g. self=bottom-right,
        other_cell=top-left, etc.

        Raises |ValueError| if the specified range already contains merged cells anywhere within
        its extents or if `other_cell` is not in the same table as `self`.
        """
        tc_range = TcRange(self._tc, other_cell._tc)

        if not tc_range.in_same_table:
            raise ValueError("other_cell from different table")
        if tc_range.contains_merged_cell:
            raise ValueError("range contains one or more merged cells")

        tc_range.move_content_to_origin()

        row_count, col_count = tc_range.dimensions

        for tc in tc_range.iter_top_row_tcs():
            tc.rowSpan = row_count
        for tc in tc_range.iter_left_col_tcs():
            tc.gridSpan = col_count
        for tc in tc_range.iter_except_left_col_tcs():
            tc.hMerge = True
        for tc in tc_range.iter_except_top_row_tcs():
            tc.vMerge = True

    @property
    def row_idx(self) -> int:
        """Zero-based row index of this cell in its table.

        Reflects the position of this cell's parent `a:tr` element among its siblings in the
        containing `a:tbl` table element. Read-only.
        """
        return self._tc.row_idx

    @property
    def span_height(self) -> int:
        """int count of rows spanned by this cell.

        The value of this property may be misleading (often 1) on cells where `.is_merge_origin`
        is not |True|, since only a merge-origin cell contains complete span information. This
        property is only intended for use on cells known to be a merge origin by testing
        `.is_merge_origin`.
        """
        return self._tc.rowSpan

    @property
    def span_width(self) -> int:
        """int count of columns spanned by this cell.

        The value of this property may be misleading (often 1) on cells where `.is_merge_origin`
        is not |True|, since only a merge-origin cell contains complete span information. This
        property is only intended for use on cells known to be a merge origin by testing
        `.is_merge_origin`.
        """
        return self._tc.gridSpan

    def split(self) -> None:
        """Remove merge from this (merge-origin) cell.

        The merged cell represented by this object will be "unmerged", yielding a separate
        unmerged cell for each grid cell previously spanned by this merge.

        Raises |ValueError| when this cell is not a merge-origin cell. Test with
        `.is_merge_origin` before calling.
        """
        if not self.is_merge_origin:
            raise ValueError("not a merge-origin cell; only a merge-origin cell can be sp" "lit")

        tc_range = TcRange.from_merge_origin(self._tc)

        for tc in tc_range.iter_tcs():
            tc.rowSpan = tc.gridSpan = 1
            tc.hMerge = tc.vMerge = False

    @property
    def text(self) -> str:
        """Textual content of cell as a single string.

        The returned string will contain a newline character (`"\\n"`) separating each paragraph
        and a vertical-tab (`"\\v"`) character for each line break (soft carriage return) in the
        cell's text.

        Assignment to `text` replaces all text currently contained in the cell. A newline
        character (`"\\n"`) in the assigned text causes a new paragraph to be started. A
        vertical-tab (`"\\v"`) character in the assigned text causes a line-break (soft
        carriage-return) to be inserted. (The vertical-tab character appears in clipboard text
        copied from PowerPoint as its encoding of line-breaks.)
        """
        return self.text_frame.text

    @text.setter
    def text(self, text: str):
        self.text_frame.text = text

    @property
    def text_frame(self) -> TextFrame:
        """|TextFrame| containing the text that appears in the cell."""
        txBody = self._tc.get_or_add_txBody()
        return TextFrame(txBody, self)

    @property
    def vertical_anchor(self) -> MSO_VERTICAL_ANCHOR | None:
        """Vertical alignment of this cell.

        This value is a member of the :ref:`MsoVerticalAnchor` enumeration or |None|. A value of
        |None| indicates the cell has no explicitly applied vertical anchor setting and its
        effective value is inherited from its style-hierarchy ancestors.

        Assigning |None| to this property causes any explicitly applied vertical anchor setting to
        be cleared and inheritance of its effective value to be restored.
        """
        return self._tc.anchor

    @vertical_anchor.setter
    def vertical_anchor(self, mso_anchor_idx: MSO_VERTICAL_ANCHOR | None):
        self._tc.anchor = mso_anchor_idx

    @staticmethod
    def _validate_margin_value(margin_value: Length | None) -> None:
        """Raise ValueError if `margin_value` is not a positive integer value or |None|."""
        if not isinstance(margin_value, int) and margin_value is not None:
            tmpl = "margin value must be integer or None, got '%s'"
            raise TypeError(tmpl % margin_value)


class _CellBorder(object):
    """Adapter that exposes one of the six `a:tcPr` border children as an `a:ln`.

    A |LineFormat| instance expects its parent to provide an ``ln`` property
    and a ``get_or_add_ln()`` method. The table-cell border elements
    (``a:lnL``, ``a:lnR``, ``a:lnT``, ``a:lnB``, ``a:lnTlToBr``,
    ``a:lnBlToTr``) share the ``CT_LineProperties`` type with ``a:ln`` but
    appear under different tag names, so this shim translates the access to
    the correct descriptor on ``a:tcPr`` for the requested side.
    """

    _VALID_SIDES = ("lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr")

    def __init__(self, tc: CT_TableCell, side: str):
        if side not in self._VALID_SIDES:
            raise ValueError(
                "side must be one of %s, got %r" % (self._VALID_SIDES, side)
            )
        self._tc = tc
        self._side = side

    @property
    def ln(self) -> CT_LineProperties | None:
        """Return the `a:lnX` child of `a:tcPr`, or |None| if not present."""
        tcPr = self._tc.tcPr
        if tcPr is None:
            return None
        return getattr(tcPr, self._side)

    def get_or_add_ln(self) -> CT_LineProperties:
        """Return the `a:lnX` child, adding it (and `a:tcPr`) if needed."""
        tcPr = self._tc.get_or_add_tcPr()
        return getattr(tcPr, "get_or_add_" + self._side)()


class _Column(Subshape):
    """Table column"""

    def __init__(self, gridCol: CT_TableCol, parent: _ColumnCollection):
        super(_Column, self).__init__(parent)
        self._parent = parent
        self._gridCol = gridCol

    def delete(self) -> None:
        """Remove this column from its containing table.

        The column's `a:gridCol` element is removed from the table's `a:tblGrid`
        and the `a:tc` cell at the same column offset is removed from every
        `a:tr` row in the table. The containing graphic-frame width is reduced
        by this column's width. Subsequent use of this column object is
        undefined; most operations will raise an exception.

        Note that deleting a column whose cells participate in a horizontal-merge
        range (as either the merge-origin or a spanned cell) may leave the table
        in an inconsistent merge state. Split any merged cells spanning the
        column first (see :meth:`._Cell.split`) if merge integrity is required.
        """
        tblGrid = self._gridCol.getparent()
        col_idx = list(tblGrid).index(self._gridCol)
        tbl = tblGrid.getparent()
        # ---remove the cell at `col_idx` in every row---
        for tr in tbl.tr_lst:
            tc = tr.tc_lst[col_idx]
            tr.remove(tc)
        # ---remove the gridCol itself---
        tblGrid.remove(self._gridCol)
        self._parent.notify_width_changed()

    @property
    def width(self) -> Length:
        """Width of column in EMU."""
        return self._gridCol.w

    @width.setter
    def width(self, width: Length):
        self._gridCol.w = width
        self._parent.notify_width_changed()


class _Row(Subshape):
    """Table row"""

    def __init__(self, tr: CT_TableRow, parent: _RowCollection):
        super(_Row, self).__init__(parent)
        self._parent = parent
        self._tr = tr

    @property
    def cells(self):
        """Read-only reference to collection of cells in row.

        An individual cell is referenced using list notation, e.g. `cell = row.cells[0]`.
        """
        return _CellCollection(self._tr, self)

    def delete(self) -> None:
        """Remove this row from its containing table.

        The row's `a:tr` element is removed from its parent `a:tbl`. The containing
        graphic-frame height is reduced by this row's height. Subsequent use of this
        row object is undefined; most operations will raise an exception.

        Note that deleting a row whose cells participate in a vertical-merge range
        (as either the merge-origin or a spanned cell) may leave the table in an
        inconsistent merge state. Split any merged cells spanning the row first
        (see :meth:`._Cell.split`) if merge integrity is required.
        """
        self._tr.getparent().remove(self._tr)
        self._parent.notify_height_changed()

    @property
    def height(self) -> Length:
        """Height of row in EMU.

        This is the *authored* (minimum) row height stored in the `.pptx` file. PowerPoint
        treats it as a minimum and will silently grow the row at render time to fit its text
        content; `python-pptx` has no access to PowerPoint's layout engine and cannot compute
        that grown height. The rendered row height may therefore exceed this value. See
        issue #296 and the "Table height and row height" section of the user guide.
        """
        return self._tr.h

    @height.setter
    def height(self, height: Length):
        self._tr.h = height
        self._parent.notify_height_changed()


class _CellCollection(Subshape):
    """Horizontal sequence of row cells"""

    def __init__(self, tr: CT_TableRow, parent: _Row):
        super(_CellCollection, self).__init__(parent)
        self._parent = parent
        self._tr = tr

    def __getitem__(self, idx: int) -> _Cell:
        """Provides indexed access, (e.g. 'cells[0]')."""
        if idx < 0 or idx >= len(self._tr.tc_lst):
            msg = "cell index [%d] out of range" % idx
            raise IndexError(msg)
        return _Cell(self._tr.tc_lst[idx], self)

    def __iter__(self) -> Iterator[_Cell]:
        """Provides iterability."""
        return (_Cell(tc, self) for tc in self._tr.tc_lst)

    def __len__(self) -> int:
        """Supports len() function (e.g. 'len(cells) == 1')."""
        return len(self._tr.tc_lst)


class _ColumnCollection(Subshape):
    """Sequence of table columns."""

    def __init__(self, tbl: CT_Table, parent: Table):
        super(_ColumnCollection, self).__init__(parent)
        self._parent = parent
        self._tbl = tbl

    def __getitem__(self, idx: int):
        """Provides indexed access, (e.g. 'columns[0]')."""
        if idx < 0 or idx >= len(self._tbl.tblGrid.gridCol_lst):
            msg = "column index [%d] out of range" % idx
            raise IndexError(msg)
        return _Column(self._tbl.tblGrid.gridCol_lst[idx], self)

    def __len__(self):
        """Supports len() function (e.g. 'len(columns) == 1')."""
        return len(self._tbl.tblGrid.gridCol_lst)

    def add(self, width: Length | None = None) -> _Column:
        """Return a newly added |_Column| appended to the right of this table.

        The new column is represented by a new `a:gridCol` child of the table's
        `a:tblGrid` element and a new `a:tc` cell appended to every existing
        `a:tr` row in the table, each containing a single empty paragraph.
        `width` is the column width in EMU. When `width` is |None| (the
        default), the new column inherits its width from the last existing
        column in the table, or defaults to 914,400 EMU (exactly 1 inch) when
        the table has no existing columns.

        Adding a column increases the width of the containing graphic-frame
        shape by the width of the new column.
        """
        if width is None:
            gridCol_lst = self._tbl.tblGrid.gridCol_lst
            width = Emu(gridCol_lst[-1].w) if gridCol_lst else Emu(914400)
        self._tbl.tblGrid.add_gridCol(width=width)
        # ---append one cell per row---
        for tr in self._tbl.tr_lst:
            tr.add_tc()
        self._parent.notify_width_changed()
        return _Column(self._tbl.tblGrid.gridCol_lst[-1], self)

    def remove(self, column: _Column) -> None:
        """Remove `column` from this table.

        `column` must be a |_Column| object belonging to this table; a
        |ValueError| is raised if it belongs to a different table. The
        column's `a:gridCol` element is detached from the table's `a:tblGrid`,
        the `a:tc` cell at the same column offset is removed from every row,
        and the graphic-frame width is reduced by the deleted column's width.

        This is the collection-level counterpart to :meth:`._Column.delete`.
        See that method for notes on merged cells crossing the deleted column.
        """
        if column._gridCol.getparent() is not self._tbl.tblGrid:
            raise ValueError("column is not a member of this table")
        col_idx = self._tbl.tblGrid.gridCol_lst.index(column._gridCol)
        # ---remove the cell at `col_idx` in every row---
        for tr in self._tbl.tr_lst:
            tc = tr.tc_lst[col_idx]
            tr.remove(tc)
        # ---remove the gridCol itself---
        self._tbl.tblGrid.remove(column._gridCol)
        self._parent.notify_width_changed()

    def notify_width_changed(self):
        """Called by a column when its width changes. Pass along to parent."""
        self._parent.notify_width_changed()


class _RowCollection(Subshape):
    """Sequence of table rows"""

    def __init__(self, tbl: CT_Table, parent: Table):
        super(_RowCollection, self).__init__(parent)
        self._parent = parent
        self._tbl = tbl

    def __getitem__(self, idx: int) -> _Row:
        """Provides indexed access, (e.g. 'rows[0]')."""
        if idx < 0 or idx >= len(self):
            msg = "row index [%d] out of range" % idx
            raise IndexError(msg)
        return _Row(self._tbl.tr_lst[idx], self)

    def __len__(self):
        """Supports len() function (e.g. 'len(rows) == 1')."""
        return len(self._tbl.tr_lst)

    def add(self, height: Length | None = None) -> _Row:
        """Return a newly added |_Row| appended to the bottom of this table.

        The new row has the same number of cells as the existing rows in the table,
        each containing a single empty paragraph. `height` is the row height in EMU.
        When `height` is |None| (the default), the new row inherits its height from
        the last existing row in the table, or defaults to 370,840 EMU (approximately
        0.4 inches) when the table has no existing rows.

        Adding a row increases the height of the containing graphic-frame shape by
        the height of the new row.
        """
        if height is None:
            tr_lst = self._tbl.tr_lst
            height = Emu(tr_lst[-1].h) if tr_lst else Emu(370840)
        tr = self._tbl.add_tr(height=height)
        # ---populate the new row with one cell per column---
        for _ in range(len(self._tbl.tblGrid.gridCol_lst)):
            tr.add_tc()
        self._parent.notify_height_changed()
        return _Row(tr, self)

    def remove(self, row: _Row) -> None:
        """Remove `row` from this table.

        `row` must be a |_Row| object belonging to this table; a |ValueError| is raised
        if it belongs to a different table. The row's `a:tr` element is detached from
        the containing `a:tbl` and the graphic-frame height is reduced by the deleted
        row's height.

        This is the collection-level counterpart to :meth:`._Row.delete`. See that
        method for notes on merged cells crossing the deleted row.
        """
        if row._tr.getparent() is not self._tbl:
            raise ValueError("row is not a member of this table")
        self._tbl.remove(row._tr)
        self._parent.notify_height_changed()

    def notify_height_changed(self):
        """Called by a row when its height changes. Pass along to parent."""
        self._parent.notify_height_changed()
