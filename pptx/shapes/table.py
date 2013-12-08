# encoding: utf-8

"""
Table-related objects such as Table and Cell.
"""

from pptx.constants import MSO
from pptx.oxml.ns import qn
from pptx.shapes import Subshape
from pptx.shapes.shape import BaseShape
from pptx.text import TextFrame
from pptx.util import to_unicode


class Table(BaseShape):
    """
    A table shape. Not intended to be constructed directly, use
    :meth:`ShapeCollection.add_table` to add a table to a slide.
    """
    def __init__(self, graphicFrame, parent):
        super(Table, self).__init__(graphicFrame, parent)
        self._graphicFrame = graphicFrame
        self._tbl_elm = graphicFrame[qn('a:graphic')].graphicData.tbl
        self._rows = _RowCollection(self._tbl_elm, self)
        self._columns = _ColumnCollection(self._tbl_elm, self)

    def cell(self, row_idx, col_idx):
        """Return table cell at *row_idx*, *col_idx* location"""
        row = self.rows[row_idx]
        return row.cells[col_idx]

    @property
    def columns(self):
        """
        Read-only reference to collection of |_Column| objects representing
        the table's columns. |_Column| objects are accessed using list
        notation, e.g. ``col = tbl.columns[0]``.
        """
        return self._columns

    @property
    def first_col(self):
        """
        Read/write boolean property which, when true, indicates the first
        column should be formatted differently, as for a side-heading column
        at the far left of the table.
        """
        return self._tbl_elm.firstCol

    @property
    def first_row(self):
        """
        Read/write boolean property which, when true, indicates the first
        row should be formatted differently, e.g. for column headings.
        """
        return self._tbl_elm.firstRow

    @property
    def height(self):
        """
        Read-only integer height of table in English Metric Units (EMU)
        """
        return int(self._graphicFrame.xfrm[qn('a:ext')].get('cy'))

    @property
    def horz_banding(self):
        """
        Read/write boolean property which, when true, indicates the rows of
        the table should appear with alternating shading.
        """
        return self._tbl_elm.bandRow

    @property
    def last_col(self):
        """
        Read/write boolean property which, when true, indicates the last
        column should be formatted differently, as for a row totals column at
        the far right of the table.
        """
        return self._tbl_elm.lastCol

    @property
    def last_row(self):
        """
        Read/write boolean property which, when true, indicates the last
        row should be formatted differently, as for a totals row at the
        bottom of the table.
        """
        return self._tbl_elm.lastRow

    @first_col.setter
    def first_col(self, value):
        self._tbl_elm.firstCol = bool(value)

    @first_row.setter
    def first_row(self, value):
        self._tbl_elm.firstRow = bool(value)

    @horz_banding.setter
    def horz_banding(self, value):
        self._tbl_elm.bandRow = bool(value)

    @last_col.setter
    def last_col(self, value):
        self._tbl_elm.lastCol = bool(value)

    def notify_height_changed(self):
        """
        Called by a row when its height changes, triggering the graphic frame
        to recalculate its total height (as the sum of the row heights).
        """
        new_table_height = sum([row.height for row in self.rows])
        self._graphicFrame.xfrm[qn('a:ext')].set('cy', str(new_table_height))

    def notify_width_changed(self):
        """
        Called by a column when its width changes, triggering the graphic
        frame to recalculate its total width (as the sum of the column
        widths).
        """
        new_table_width = sum([col.width for col in self.columns])
        self._graphicFrame.xfrm[qn('a:ext')].set('cx', str(new_table_width))

    @last_row.setter
    def last_row(self, value):
        self._tbl_elm.lastRow = bool(value)

    @property
    def rows(self):
        """
        Read-only reference to collection of |_Row| objects representing the
        table's rows. |_Row| objects are accessed using list notation, e.g.
        ``col = tbl.rows[0]``.
        """
        return self._rows

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, unconditionally
        ``MSO.TABLE`` in this case.
        """
        return MSO.TABLE

    @property
    def vert_banding(self):
        """
        Read/write boolean property which, when true, indicates the columns
        of the table should appear with alternating shading.
        """
        return self._tbl_elm.bandCol

    @vert_banding.setter
    def vert_banding(self, value):
        self._tbl_elm.bandCol = bool(value)

    @property
    def width(self):
        """
        Read-only integer width of table in English Metric Units (EMU)
        """
        return int(self._graphicFrame.xfrm[qn('a:ext')].get('cx'))


class _Cell(Subshape):
    """
    Table cell
    """
    def __init__(self, tc, parent):
        super(_Cell, self).__init__(parent)
        self._tc = tc

    @property
    def margin_top(self):
        """
        Read/write integer value of top margin of cell in English Metric
        Units (EMU). If assigned |None|, the default value is used, 0.1
        inches for left and right margins and 0.05 inches for top and bottom.
        """
        return self._tc.marT

    @property
    def margin_right(self):
        """Right margin of cell"""
        return self._tc.marR

    @property
    def margin_bottom(self):
        """Bottom margin of cell"""
        return self._tc.marB

    @property
    def margin_left(self):
        """Left margin of cell"""
        return self._tc.marL

    @margin_top.setter
    def margin_top(self, margin_top):
        self._validate_margin_value(margin_top)
        self._tc.marT = margin_top

    @margin_right.setter
    def margin_right(self, margin_right):
        self._validate_margin_value(margin_right)
        self._tc.marR = margin_right

    @margin_bottom.setter
    def margin_bottom(self, margin_bottom):
        self._validate_margin_value(margin_bottom)
        self._tc.marB = margin_bottom

    @margin_left.setter
    def margin_left(self, margin_left):
        self._validate_margin_value(margin_left)
        self._tc.marL = margin_left

    def text(self, text):
        """
        Replace all text in cell with single run containing *text*
        """
        self.textframe.text = to_unicode(text)

    #: Write-only. Assignment to *text* replaces all text currently contained
    #: in the cell, resulting in a text frame containing exactly one
    #: paragraph, itself containing a single run. The assigned value can be a
    #: 7-bit ASCII string, a UTF-8 encoded 8-bit string, or unicode. String
    #: values are converted to unicode assuming UTF-8 encoding.
    text = property(None, text)

    @property
    def textframe(self):
        """
        |TextFrame| instance containing the text that appears in the cell.
        """
        txBody = self._tc.get_or_add_txBody()
        return TextFrame(txBody, self)

    @property
    def vertical_anchor(self):
        """
        Vertical anchor of this table cell, determines the vertical alignment
        of text in the cell. Value is like ``MSO.ANCHOR_MIDDLE``. Can be
        |None|, meaning the cell has no vertical anchor setting and its
        effective value is inherited from a higher-level object.
        """
        return self._tc.anchor

    @vertical_anchor.setter
    def vertical_anchor(self, mso_anchor_idx):
        """
        Set vertical_anchor of this cell to *vertical_anchor*, a constant
        value like ``MSO_ANCHOR.MIDDLE``. If *vertical_anchor* is |None|, any
        vertical anchor setting is cleared and its effective value is
        inherited.
        """
        self._tc.anchor = mso_anchor_idx

    @staticmethod
    def _validate_margin_value(margin_value):
        """
        Raise ValueError if *margin_value* is not a positive integer value or
        |None|.
        """
        if (not isinstance(margin_value, (int, long))
                and margin_value is not None):
            tmpl = "margin value must be integer or None, got '%s'"
            raise TypeError(tmpl % margin_value)


class _Column(Subshape):
    """
    Table column
    """
    def __init__(self, gridCol, parent):
        super(_Column, self).__init__(parent)
        self._gridCol = gridCol

    def _get_width(self):
        """
        Return width of column in EMU
        """
        return int(self._gridCol.get('w'))

    def _set_width(self, width):
        """
        Set column width to *width*, a positive integer value.
        """
        if not isinstance(width, int) or width < 0:
            msg = "column width must be positive integer"
            raise ValueError(msg)
        self._gridCol.set('w', str(width))
        self._parent.notify_width_changed()

    #: Read-write integer width of this column in English Metric Units (EMU).
    width = property(_get_width, _set_width)


class _Row(Subshape):
    """
    Table row
    """
    def __init__(self, tr, parent):
        super(_Row, self).__init__(parent)
        self._tr = tr
        self._cells = _CellCollection(tr, self)

    @property
    def cells(self):
        """
        Read-only reference to collection of cells in row. An individual cell
        is referenced using list notation, e.g. ``cell = row.cells[0]``.
        """
        return self._cells

    def _get_height(self):
        """
        Return height of row in EMU
        """
        return int(self._tr.get('h'))

    def _set_height(self, height):
        """
        Set row height to *height*, a positive integer value.
        """
        if not isinstance(height, int) or height < 0:
            msg = "row height must be positive integer"
            raise ValueError(msg)
        self._tr.set('h', str(height))
        self._parent.notify_height_changed()

    #: Read/write integer height of this row in English Metric Units (EMU).
    height = property(_get_height, _set_height)


class _CellCollection(Subshape):
    """
    "Horizontal" sequence of row cells
    """
    def __init__(self, tr, parent):
        super(_CellCollection, self).__init__(parent)
        self._tr = tr

    def __getitem__(self, idx):
        """Provides indexed access, (e.g. 'cells[0]')."""
        if idx < 0 or idx >= len(self._tr.tc):
            msg = "cell index [%d] out of range" % idx
            raise IndexError(msg)
        return _Cell(self._tr.tc[idx], self)

    def __len__(self):
        """Supports len() function (e.g. 'len(cells) == 1')."""
        return len(self._tr.tc)


class _ColumnCollection(Subshape):
    """
    Sequence of table columns.
    """
    def __init__(self, tbl_elm, parent):
        super(_ColumnCollection, self).__init__(parent)
        self._tbl_elm = tbl_elm

    def __getitem__(self, idx):
        """Provides indexed access, (e.g. 'columns[0]')."""
        if idx < 0 or idx >= len(self._tbl_elm.tblGrid.gridCol):
            msg = "column index [%d] out of range" % idx
            raise IndexError(msg)
        return _Column(self._tbl_elm.tblGrid.gridCol[idx], self)

    def __len__(self):
        """Supports len() function (e.g. 'len(columns) == 1')."""
        return len(self._tbl_elm.tblGrid.gridCol)

    def notify_width_changed(self):
        """
        Called by a column when its width changes. Pass along to parent.
        """
        self._parent.notify_width_changed()


class _RowCollection(Subshape):
    """
    Sequence of table rows.
    """
    def __init__(self, tbl_elm, parent):
        super(_RowCollection, self).__init__(parent)
        self._tbl_elm = tbl_elm

    def __getitem__(self, idx):
        """Provides indexed access, (e.g. 'rows[0]')."""
        if idx < 0 or idx >= len(self._tbl_elm.tr):
            msg = "row index [%d] out of range" % idx
            raise IndexError(msg)
        return _Row(self._tbl_elm.tr[idx], self)

    def __len__(self):
        """Supports len() function (e.g. 'len(rows) == 1')."""
        return len(self._tbl_elm.tr)

    def notify_height_changed(self):
        """
        Called by a row when its height changes. Pass along to parent.
        """
        self._parent.notify_height_changed()
