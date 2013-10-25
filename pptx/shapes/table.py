# encoding: utf-8

"""
Table-related objects such as Table and Cell.
"""

from pptx.constants import MSO
from pptx.oxml import qn
from pptx.shapes.shape import _BaseShape
from pptx.spec import namespaces, VerticalAnchor
from pptx.text import _TextFrame

# default namespace map for use in lxml calls
_nsmap = namespaces('a', 'r', 'p')


def _child(element, child_tagname, nsmap=None):
    """
    Return direct child of *element* having *child_tagname* or |None| if no
    such child element is present.
    """
    # use default nsmap if not specified
    if nsmap is None:
        nsmap = _nsmap
    xpath = './%s' % child_tagname
    matching_children = element.xpath(xpath, namespaces=nsmap)
    return matching_children[0] if len(matching_children) else None


def _to_unicode(text):
    """
    Return *text* as a unicode string.

    *text* can be a 7-bit ASCII string, a UTF-8 encoded 8-bit string, or
    unicode. String values are converted to unicode assuming UTF-8 encoding.
    Unicode values are returned unchanged.
    """
    # both str and unicode inherit from basestring
    if not isinstance(text, basestring):
        tmpl = 'expected UTF-8 encoded string or unicode, got %s value %s'
        raise TypeError(tmpl % (type(text), text))
    # return unicode strings unchanged
    if isinstance(text, unicode):
        return text
    # otherwise assume UTF-8 encoding, which also works for ASCII
    return unicode(text, 'utf-8')


class _Table(_BaseShape):
    """
    A table shape. Not intended to be constructed directly, use
    :meth:`_ShapeCollection.add_table` to add a table to a slide.
    """
    def __init__(self, graphicFrame):
        super(_Table, self).__init__(graphicFrame)
        self.__graphicFrame = graphicFrame
        self.__tbl_elm = graphicFrame[qn('a:graphic')].graphicData.tbl
        self.__rows = _RowCollection(self.__tbl_elm, self)
        self.__columns = _ColumnCollection(self.__tbl_elm, self)

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
        return self.__columns

    @property
    def first_col(self):
        """
        Read/write boolean property which, when true, indicates the first
        column should be formatted differently, as for a side-heading column
        at the far left of the table.
        """
        return self.__tbl_elm.firstCol

    @property
    def first_row(self):
        """
        Read/write boolean property which, when true, indicates the first
        row should be formatted differently, e.g. for column headings.
        """
        return self.__tbl_elm.firstRow

    @property
    def horz_banding(self):
        """
        Read/write boolean property which, when true, indicates the rows of
        the table should appear with alternating shading.
        """
        return self.__tbl_elm.bandRow

    @property
    def last_col(self):
        """
        Read/write boolean property which, when true, indicates the last
        column should be formatted differently, as for a row totals column at
        the far right of the table.
        """
        return self.__tbl_elm.lastCol

    @property
    def last_row(self):
        """
        Read/write boolean property which, when true, indicates the last
        row should be formatted differently, as for a totals row at the
        bottom of the table.
        """
        return self.__tbl_elm.lastRow

    @property
    def vert_banding(self):
        """
        Read/write boolean property which, when true, indicates the columns
        of the table should appear with alternating shading.
        """
        return self.__tbl_elm.bandCol

    @first_col.setter
    def first_col(self, value):
        self.__tbl_elm.firstCol = bool(value)

    @first_row.setter
    def first_row(self, value):
        self.__tbl_elm.firstRow = bool(value)

    @horz_banding.setter
    def horz_banding(self, value):
        self.__tbl_elm.bandRow = bool(value)

    @last_col.setter
    def last_col(self, value):
        self.__tbl_elm.lastCol = bool(value)

    @last_row.setter
    def last_row(self, value):
        self.__tbl_elm.lastRow = bool(value)

    @vert_banding.setter
    def vert_banding(self, value):
        self.__tbl_elm.bandCol = bool(value)

    @property
    def height(self):
        """
        Read-only integer height of table in English Metric Units (EMU)
        """
        return int(self.__graphicFrame.xfrm[qn('a:ext')].get('cy'))

    @property
    def rows(self):
        """
        Read-only reference to collection of |_Row| objects representing the
        table's rows. |_Row| objects are accessed using list notation, e.g.
        ``col = tbl.rows[0]``.
        """
        return self.__rows

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, unconditionally
        ``MSO.TABLE`` in this case.
        """
        return MSO.TABLE

    @property
    def width(self):
        """
        Read-only integer width of table in English Metric Units (EMU)
        """
        return int(self.__graphicFrame.xfrm[qn('a:ext')].get('cx'))

    def _notify_height_changed(self):
        """
        Called by a row when its height changes, triggering the graphic frame
        to recalculate its total height (as the sum of the row heights).
        """
        new_table_height = sum([row.height for row in self.rows])
        self.__graphicFrame.xfrm[qn('a:ext')].set('cy', str(new_table_height))

    def _notify_width_changed(self):
        """
        Called by a column when its width changes, triggering the graphic
        frame to recalculate its total width (as the sum of the column
        widths).
        """
        new_table_width = sum([col.width for col in self.columns])
        self.__graphicFrame.xfrm[qn('a:ext')].set('cx', str(new_table_width))


class _Cell(object):
    """
    Table cell
    """

    def __init__(self, tc):
        super(_Cell, self).__init__()
        self.__tc = tc

    @staticmethod
    def __assert_valid_margin_value(margin_value):
        """
        Raise ValueError if *margin_value* is not a positive integer value or
        |None|.
        """
        if (not isinstance(margin_value, (int, long))
                and margin_value is not None):
            tmpl = "margin value must be integer or None, got '%s'"
            raise ValueError(tmpl % margin_value)

    @property
    def margin_top(self):
        """
        Read/write integer value of top margin of cell in English Metric
        Units (EMU). If assigned |None|, the default value is used, 0.1
        inches for left and right margins and 0.05 inches for top and bottom.
        """
        return self.__tc.marT

    @property
    def margin_right(self):
        """Right margin of cell"""
        return self.__tc.marR

    @property
    def margin_bottom(self):
        """Bottom margin of cell"""
        return self.__tc.marB

    @property
    def margin_left(self):
        """Left margin of cell"""
        return self.__tc.marL

    @margin_top.setter
    def margin_top(self, margin_top):
        self.__assert_valid_margin_value(margin_top)
        self.__tc.marT = margin_top

    @margin_right.setter
    def margin_right(self, margin_right):
        self.__assert_valid_margin_value(margin_right)
        self.__tc.marR = margin_right

    @margin_bottom.setter
    def margin_bottom(self, margin_bottom):
        self.__assert_valid_margin_value(margin_bottom)
        self.__tc.marB = margin_bottom

    @margin_left.setter
    def margin_left(self, margin_left):
        self.__assert_valid_margin_value(margin_left)
        self.__tc.marL = margin_left

    def _set_text(self, text):
        """Replace all text in cell with single run containing *text*"""
        self.textframe.text = _to_unicode(text)

    #: Write-only. Assignment to *text* replaces all text currently contained
    #: in the cell, resulting in a text frame containing exactly one
    #: paragraph, itself containing a single run. The assigned value can be a
    #: 7-bit ASCII string, a UTF-8 encoded 8-bit string, or unicode. String
    #: values are converted to unicode assuming UTF-8 encoding.
    text = property(None, _set_text)

    @property
    def textframe(self):
        """
        |_TextFrame| instance containing the text that appears in the cell.
        """
        txBody = _child(self.__tc, 'a:txBody')
        if txBody is None:
            raise ValueError('cell has no text frame')
        return _TextFrame(txBody)

    def _get_vertical_anchor(self):
        """
        Return vertical anchor setting for this table cell, e.g.
        ``MSO.ANCHOR_MIDDLE``. Can be |None|, meaning the cell has no
        vertical anchor setting and its effective value is inherited from a
        higher-level object.
        """
        anchor = self.__tc.anchor
        return VerticalAnchor.from_text_anchoring_type(anchor)

    def _set_vertical_anchor(self, vertical_anchor):
        """
        Set vertical_anchor of this cell to *vertical_anchor*, a constant
        value like ``MSO.ANCHOR_MIDDLE``. If *vertical_anchor* is |None|, any
        vertical anchor setting is cleared and its effective value is
        inherited.
        """
        anchor = VerticalAnchor.to_text_anchoring_type(vertical_anchor)
        self.__tc.anchor = anchor

    #: Vertical anchor of this table cell, determines the vertical alignment
    #: of text in the cell. Value is like ``MSO.ANCHOR_MIDDLE``. Can be
    #: |None|, meaning the cell has no vertical anchor setting and its
    #: effective value is inherited from a higher-level object.
    vertical_anchor = property(_get_vertical_anchor, _set_vertical_anchor)


class _Column(object):
    """
    Table column
    """

    def __init__(self, gridCol, table):
        super(_Column, self).__init__()
        self.__gridCol = gridCol
        self.__table = table

    def _get_width(self):
        """
        Return width of column in EMU
        """
        return int(self.__gridCol.get('w'))

    def _set_width(self, width):
        """
        Set column width to *width*, a positive integer value.
        """
        if not isinstance(width, int) or width < 0:
            msg = "column width must be positive integer"
            raise ValueError(msg)
        self.__gridCol.set('w', str(width))
        self.__table._notify_width_changed()

    #: Read-write integer width of this column in English Metric Units (EMU).
    width = property(_get_width, _set_width)


class _Row(object):
    """
    Table row
    """

    def __init__(self, tr, table):
        super(_Row, self).__init__()
        self.__tr = tr
        self.__table = table
        self.__cells = _CellCollection(tr)

    @property
    def cells(self):
        """
        Read-only reference to collection of cells in row. An individual cell
        is referenced using list notation, e.g. ``cell = row.cells[0]``.
        """
        return self.__cells

    def _get_height(self):
        """
        Return height of row in EMU
        """
        return int(self.__tr.get('h'))

    def _set_height(self, height):
        """
        Set row height to *height*, a positive integer value.
        """
        if not isinstance(height, int) or height < 0:
            msg = "row height must be positive integer"
            raise ValueError(msg)
        self.__tr.set('h', str(height))
        self.__table._notify_height_changed()

    #: Read/write integer height of this row in English Metric Units (EMU).
    height = property(_get_height, _set_height)


class _CellCollection(object):
    """
    "Horizontal" sequence of row cells
    """

    def __init__(self, tr):
        super(_CellCollection, self).__init__()
        self.__tr = tr

    def __getitem__(self, idx):
        """Provides indexed access, (e.g. 'cells[0]')."""
        if idx < 0 or idx >= len(self.__tr.tc):
            msg = "cell index [%d] out of range" % idx
            raise IndexError(msg)
        return _Cell(self.__tr.tc[idx])

    def __len__(self):
        """Supports len() function (e.g. 'len(cells) == 1')."""
        return len(self.__tr.tc)


class _ColumnCollection(object):
    """
    Sequence of table columns.
    """

    def __init__(self, tbl_elm, table):
        super(_ColumnCollection, self).__init__()
        self.__tbl_elm = tbl_elm
        self.__table = table

    def __getitem__(self, idx):
        """Provides indexed access, (e.g. 'columns[0]')."""
        if idx < 0 or idx >= len(self.__tbl_elm.tblGrid.gridCol):
            msg = "column index [%d] out of range" % idx
            raise IndexError(msg)
        return _Column(self.__tbl_elm.tblGrid.gridCol[idx], self.__table)

    def __len__(self):
        """Supports len() function (e.g. 'len(columns) == 1')."""
        return len(self.__tbl_elm.tblGrid.gridCol)


class _RowCollection(object):
    """
    Sequence of table rows.
    """

    def __init__(self, tbl_elm, table):
        super(_RowCollection, self).__init__()
        self.__tbl_elm = tbl_elm
        self.__table = table

    def __getitem__(self, idx):
        """Provides indexed access, (e.g. 'rows[0]')."""
        if idx < 0 or idx >= len(self.__tbl_elm.tr):
            msg = "row index [%d] out of range" % idx
            raise IndexError(msg)
        return _Row(self.__tbl_elm.tr[idx], self.__table)

    def __len__(self):
        """Supports len() function (e.g. 'len(rows) == 1')."""
        return len(self.__tbl_elm.tr)
