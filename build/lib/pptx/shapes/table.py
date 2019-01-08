# encoding: utf-8

"""
Table-related objects such as Table and Cell.
"""

from __future__ import absolute_import, print_function

from . import Subshape
from ..compat import is_integer, to_unicode
from ..dml.fill import FillFormat
from ..text.text import TextFrame
from ..util import lazyproperty


class Table(object):
    """
    A table shape. Not intended to be constructed directly, use
    :meth:`.Slide.shapes.add_table` to add a table to a slide.
    """
    def __init__(self, tbl, graphic_frame):
        super(Table, self).__init__()
        self._tbl = tbl
        self._graphic_frame = graphic_frame

    def cell(self, row_idx, col_idx):
        """
        Return table cell at *row_idx*, *col_idx* location. Indexes are
        zero-based, e.g. cell(0, 0) is the top, left cell.
        """
        row = self.rows[row_idx]
        return row.cells[col_idx]

    @lazyproperty
    def columns(self):
        """
        Read-only reference to collection of |_Column| objects representing
        the table's columns. |_Column| objects are accessed using list
        notation, e.g. ``col = tbl.columns[0]``.
        """
        return _ColumnCollection(self._tbl, self)

    @property
    def first_col(self):
        """
        Read/write boolean property which, when true, indicates the first
        column should be formatted differently, as for a side-heading column
        at the far left of the table.
        """
        return self._tbl.firstCol

    @first_col.setter
    def first_col(self, value):
        self._tbl.firstCol = value

    @property
    def first_row(self):
        """
        Read/write boolean property which, when true, indicates the first
        row should be formatted differently, e.g. for column headings.
        """
        return self._tbl.firstRow

    @first_row.setter
    def first_row(self, value):
        self._tbl.firstRow = value

    @property
    def horz_banding(self):
        """
        Read/write boolean property which, when true, indicates the rows of
        the table should appear with alternating shading.
        """
        return self._tbl.bandRow

    @horz_banding.setter
    def horz_banding(self, value):
        self._tbl.bandRow = value

    @property
    def last_col(self):
        """
        Read/write boolean property which, when true, indicates the last
        column should be formatted differently, as for a row totals column at
        the far right of the table.
        """
        return self._tbl.lastCol

    @last_col.setter
    def last_col(self, value):
        self._tbl.lastCol = value

    @property
    def last_row(self):
        """
        Read/write boolean property which, when true, indicates the last
        row should be formatted differently, as for a totals row at the
        bottom of the table.
        """
        return self._tbl.lastRow

    @last_row.setter
    def last_row(self, value):
        self._tbl.lastRow = value

    def notify_height_changed(self):
        """
        Called by a row when its height changes, triggering the graphic frame
        to recalculate its total height (as the sum of the row heights).
        """
        new_table_height = sum([row.height for row in self.rows])
        self._graphic_frame.height = new_table_height

    def notify_width_changed(self):
        """
        Called by a column when its width changes, triggering the graphic
        frame to recalculate its total width (as the sum of the column
        widths).
        """
        new_table_width = sum([col.width for col in self.columns])
        self._graphic_frame.width = new_table_width

    @property
    def part(self):
        """
        The package part containing this table.
        """
        return self._graphic_frame.part

    @lazyproperty
    def rows(self):
        """
        Read-only reference to collection of |_Row| objects representing the
        table's rows. |_Row| objects are accessed using list notation, e.g.
        ``col = tbl.rows[0]``.
        """
        return _RowCollection(self._tbl, self)

    @property
    def vert_banding(self):
        """
        Read/write boolean property which, when true, indicates the columns
        of the table should appear with alternating shading.
        """
        return self._tbl.bandCol

    @vert_banding.setter
    def vert_banding(self, value):
        self._tbl.bandCol = value


class _Cell(Subshape):
    """
    Table cell
    """
    def __init__(self, tc, parent):
        super(_Cell, self).__init__(parent)
        self._tc = tc

    @lazyproperty
    def fill(self):
        """
        |FillFormat| instance for this cell, providing access to fill
        properties such as foreground color.
        """
        tcPr = self._tc.get_or_add_tcPr()
        return FillFormat.from_fill_parent(tcPr)

    @property
    def margin_left(self):
        """
        Read/write integer value of left margin of cell as a |Length| value
        object. If assigned |None|, the default value is used, 0.1 inches for
        left and right margins and 0.05 inches for top and bottom.
        """
        return self._tc.marL

    @margin_left.setter
    def margin_left(self, margin_left):
        self._validate_margin_value(margin_left)
        self._tc.marL = margin_left

    @property
    def margin_right(self):
        """
        Right margin of cell.
        """
        return self._tc.marR

    @margin_right.setter
    def margin_right(self, margin_right):
        self._validate_margin_value(margin_right)
        self._tc.marR = margin_right

    @property
    def margin_top(self):
        """
        Top margin of cell.
        """
        return self._tc.marT

    @margin_top.setter
    def margin_top(self, margin_top):
        self._validate_margin_value(margin_top)
        self._tc.marT = margin_top

    @property
    def margin_bottom(self):
        """
        Bottom margin of cell.
        """
        return self._tc.marB

    @margin_bottom.setter
    def margin_bottom(self, margin_bottom):
        self._validate_margin_value(margin_bottom)
        self._tc.marB = margin_bottom

    def text(self, text):
        """
        Replace all text in cell with single run containing *text*
        """
        self.text_frame.text = to_unicode(text)

    #: Write-only. Assignment to *text* replaces all text currently contained
    #: in the cell, resulting in a text frame containing exactly one
    #: paragraph, itself containing a single run. The assigned value can be a
    #: 7-bit ASCII string, a UTF-8 encoded 8-bit string, or unicode. String
    #: values are converted to unicode assuming UTF-8 encoding.
    text = property(None, text)

    @property
    def text_frame(self):
        """
        |TextFrame| instance containing the text that appears in the cell.
        """
        txBody = self._tc.get_or_add_txBody()
        return TextFrame(txBody, self)

    @property
    def vertical_anchor(self):
        """
        Vertical anchor of this table cell, determines the vertical alignment
        of text in the cell. Value is like ``MSO_ANCHOR.MIDDLE``. Can be
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
        if (not is_integer(margin_value) and margin_value is not None):
            tmpl = "margin value must be integer or None, got '%s'"
            raise TypeError(tmpl % margin_value)


class _Column(Subshape):
    """
    Table column
    """
    def __init__(self, gridCol, parent):
        super(_Column, self).__init__(parent)
        self._gridCol = gridCol

    @property
    def width(self):
        """
        Width of column in EMU.
        """
        return self._gridCol.w

    @width.setter
    def width(self, width):
        self._gridCol.w = width
        self._parent.notify_width_changed()


class _Row(Subshape):
    """
    Table row
    """
    def __init__(self, tr, parent):
        super(_Row, self).__init__(parent)
        self._tr = tr

    @property
    def cells(self):
        """
        Read-only reference to collection of cells in row. An individual cell
        is referenced using list notation, e.g. ``cell = row.cells[0]``.
        """
        return _CellCollection(self._tr, self)

    @property
    def height(self):
        """
        Height of row in EMU.
        """
        return self._tr.h

    @height.setter
    def height(self, height):
        self._tr.h = height
        self._parent.notify_height_changed()


class _CellCollection(Subshape):
    """
    "Horizontal" sequence of row cells
    """
    def __init__(self, tr, parent):
        super(_CellCollection, self).__init__(parent)
        self._tr = tr

    def __getitem__(self, idx):
        """
        Provides indexed access, (e.g. 'cells[0]').
        """
        if idx < 0 or idx >= len(self._tr.tc_lst):
            msg = "cell index [%d] out of range" % idx
            raise IndexError(msg)
        return _Cell(self._tr.tc_lst[idx], self)

    def __len__(self):
        """
        Supports len() function (e.g. 'len(cells) == 1').
        """
        return len(self._tr.tc_lst)


class _ColumnCollection(Subshape):
    """
    Sequence of table columns.
    """
    def __init__(self, tbl, parent):
        super(_ColumnCollection, self).__init__(parent)
        self._tbl = tbl

    def __getitem__(self, idx):
        """
        Provides indexed access, (e.g. 'columns[0]').
        """
        if idx < 0 or idx >= len(self._tbl.tblGrid.gridCol_lst):
            msg = "column index [%d] out of range" % idx
            raise IndexError(msg)
        return _Column(self._tbl.tblGrid.gridCol_lst[idx], self)

    def __len__(self):
        """
        Supports len() function (e.g. 'len(columns) == 1').
        """
        return len(self._tbl.tblGrid.gridCol_lst)

    def notify_width_changed(self):
        """
        Called by a column when its width changes. Pass along to parent.
        """
        self._parent.notify_width_changed()


class _RowCollection(Subshape):
    """
    Sequence of table rows.
    """
    def __init__(self, tbl, parent):
        super(_RowCollection, self).__init__(parent)
        self._tbl = tbl

    def __getitem__(self, idx):
        """
        Provides indexed access, (e.g. 'rows[0]').
        """
        if idx < 0 or idx >= len(self):
            msg = "row index [%d] out of range" % idx
            raise IndexError(msg)
        return _Row(self._tbl.tr_lst[idx], self)

    def __len__(self):
        """
        Supports len() function (e.g. 'len(rows) == 1').
        """
        return len(self._tbl.tr_lst)

    def notify_height_changed(self):
        """
        Called by a row when its height changes. Pass along to parent.
        """
        self._parent.notify_height_changed()
