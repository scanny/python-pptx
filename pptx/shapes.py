# -*- coding: utf-8 -*-
#
# shapes.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""
Classes that implement PowerPoint shapes such as picture, textbox, and table.
"""

from pptx.constants import MSO
from pptx.oxml import (
    _get_or_add, qn, _Element, _SubElement, CT_GraphicalObjectFrame,
    CT_Picture, CT_Shape
)
from pptx.spec import (
    autoshape_types, namespaces, ParagraphAlignment, slide_ph_basenames,
    VerticalAnchor
)
from pptx.spec import (
    PH_ORIENT_HORZ, PH_ORIENT_VERT, PH_SZ_FULL, PH_TYPE_DT, PH_TYPE_FTR,
    PH_TYPE_OBJ, PH_TYPE_SLDNUM)
from pptx.util import Collection

# import logging
# log = logging.getLogger('pptx.shapes')

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


# ============================================================================
# Shapes
# ============================================================================

class _AutoShapeType(object):
    """
    Return an instance of |_AutoShapeType| containing metadata for an auto
    shape of type identified by *autoshape_type_id*. Instances are cached, so
    no more than one instance for a particular auto shape type is in memory.

    Instances provide the following attributes:

    .. attribute:: autoshape_type_id

       Integer uniquely identifying this auto shape type. Corresponds to a
       value in ``pptx.constants.MSO`` like ``MSO.SHAPE_ROUNDED_RECTANGLE``.

    .. attribute:: basename

       Base part of shape name for auto shapes of this type, e.g. ``Rounded
       Rectangle`` becomes ``Rounded Rectangle 99`` when the distinguishing
       integer is added to the shape name.

    .. attribute:: prst

       String identifier for this auto shape type used in the ``<a:prstGeom>``
       element.

    .. attribute:: desc

       Informal string description of auto shape.

    """
    __instances = {}

    def __new__(cls, autoshape_type_id):
        """
        Only create new instance on first call for content_type. After that,
        use cached instance.
        """
        # if there's not a matching instance in the cache, create one
        if autoshape_type_id not in cls.__instances:
            inst = super(_AutoShapeType, cls).__new__(cls)
            cls.__instances[autoshape_type_id] = inst
        # return the instance; note that __init__() gets called either way
        return cls.__instances[autoshape_type_id]

    def __init__(self, autoshape_type_id):
        """Initialize attributes from constant values in pptx.spec"""
        # skip loading if this instance is from the cache
        if hasattr(self, '_loaded'):
            return
        # raise on bad autoshape_type_id
        if autoshape_type_id not in autoshape_types:
            tmpl = "no autoshape type with id %d in pptx.spec.autoshape_types"
            raise KeyError(tmpl % autoshape_type_id)
        # otherwise initialize new instance
        autoshape_type = autoshape_types[autoshape_type_id]
        self.__autoshape_type_id = autoshape_type_id
        self.__prst = autoshape_type['prst']
        self.__basename = autoshape_type['basename']
        self.__desc = autoshape_type['desc']
        self._loaded = True

    @property
    def autoshape_type_id(self):
        """Integer identifier of this auto shape type"""
        return self.__autoshape_type_id

    @property
    def basename(self):
        """
        Base of shape name (less the distinguishing integer) for this auto
        shape type
        """
        return self.__basename

    @property
    def desc(self):
        """Informal description of this auto shape type"""
        return self.__desc

    @staticmethod
    def _lookup_id_by_prst(prst):
        """
        Return auto shape id (e.g. ``MSO.SHAPE_RECTANGLE``) corresponding to
        specified preset geometry keyword *prst*.
        """
        for autoshape_type_id, attribs in autoshape_types.iteritems():
            if attribs['prst'] == prst:
                return autoshape_type_id
        msg = "no auto shape with prst '%s'" % prst
        raise KeyError(msg)

    @property
    def prst(self):
        """
        Preset geometry identifier string for this auto shape. Used in the
        ``prst`` attribute of ``<a:prstGeom>`` element to specify the geometry
        to be used in rendering the shape, for example ``'roundRect'``.
        """
        return self.__prst


class _BaseShape(object):
    """
    Base class for shape objects. Both |_Shape| and |_Picture| inherit from
    |_BaseShape|.
    """
    def __init__(self, shape_element):
        super(_BaseShape, self).__init__()
        self._element = shape_element
        # e.g. nvSpPr for shape, nvPicPr for pic, etc.
        self.__nvXxPr = shape_element.xpath('./*[1]', namespaces=_nsmap)[0]

    @property
    def has_textframe(self):
        """
        True if this shape has a txBody element and can contain text.
        """
        return _child(self._element, 'p:txBody') is not None

    @property
    def id(self):
        """
        Id of this shape. Note that ids are constrained to positive integers.
        """
        return int(self.__nvXxPr.cNvPr.get('id'))

    @property
    def is_placeholder(self):
        """
        True if this shape is a placeholder. A shape is a placeholder if it
        has a <p:ph> element.
        """
        return _child(self.__nvXxPr.nvPr, 'p:ph') is not None

    @property
    def name(self):
        """Name of this shape."""
        return self.__nvXxPr.cNvPr.get('name')

    def _set_text(self, text):
        """Replace all text in shape with single run containing *text*"""
        if not self.has_textframe:
            raise TypeError("cannot set text of shape with no text frame")
        self.textframe.text = _to_unicode(text)

    #: Write-only. Assignment to *text* replaces all text currently contained
    #: by the shape, resulting in a text frame containing exactly one
    #: paragraph, itself containing a single run. The assigned value can be a
    #: 7-bit ASCII string, a UTF-8 encoded 8-bit string, or unicode. String
    #: values are converted to unicode assuming UTF-8 encoding.
    text = property(None, _set_text)

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, like ``MSO.CHART``.
        Must be implemented by subclasses. This one returns |None|
        unconditionally to account for shapes that haven't been implemented
        yet, like group shape and chart. Once those are done this should raise
        |NotImplementedError|.
        """
        return None
        # msg = 'shape_type property must be implemented by subclasses'
        # raise NotImplementedError(msg)

    @property
    def textframe(self):
        """
        _TextFrame instance for this shape. Raises |ValueError| if shape has
        no text frame. Use :attr:`has_textframe` to check whether a shape has
        a text frame.
        """
        txBody = _child(self._element, 'p:txBody')
        if txBody is None:
            raise ValueError('shape has no text frame')
        return _TextFrame(txBody)

    @property
    def _is_title(self):
        """
        True if this shape is a title placeholder.
        """
        ph = _child(self.__nvXxPr.nvPr, 'p:ph')
        if ph is None:
            return False
        # idx defaults to 0 when idx attr is absent
        ph_idx = ph.get('idx', '0')
        # title placeholder is identified by idx of 0
        return ph_idx == '0'


class _ShapeCollection(_BaseShape, Collection):
    """
    Sequence of shapes. Corresponds to CT_GroupShape in pml schema. Note that
    while spTree in a slide is a group shape, the group shape is recursive in
    that a group shape can include other group shapes within it.
    """
    _NVGRPSPPR = qn('p:nvGrpSpPr')
    _GRPSPPR = qn('p:grpSpPr')
    _SP = qn('p:sp')
    _GRPSP = qn('p:grpSp')
    _GRAPHICFRAME = qn('p:graphicFrame')
    _CXNSP = qn('p:cxnSp')
    _PIC = qn('p:pic')
    _CONTENTPART = qn('p:contentPart')
    _EXTLST = qn('p:extLst')

    def __init__(self, spTree, slide=None):
        super(_ShapeCollection, self).__init__(spTree)
        self.__spTree = spTree
        self.__slide = slide
        self.__shapes = self._values
        # unmarshal shapes
        for elm in spTree.iterchildren():
            if elm.tag in (self._NVGRPSPPR, self._GRPSPPR, self._EXTLST):
                continue
            elif elm.tag == self._SP:
                shape = _Shape(elm)
            elif elm.tag == self._PIC:
                shape = _Picture(elm)
            elif elm.tag == self._GRPSP:
                shape = _ShapeCollection(elm)
            elif elm.tag == self._GRAPHICFRAME:
                if elm.has_table:
                    shape = _Table(elm)
                else:
                    shape = _BaseShape(elm)
            elif elm.tag == self._CONTENTPART:
                msg = ("first time 'contentPart' shape encountered in the "
                       "wild, please let developer know and send example")
                raise ValueError(msg)
            else:
                shape = _BaseShape(elm)
            self.__shapes.append(shape)

    @property
    def placeholders(self):
        """
        Immutable sequence containing the placeholder shapes in this shape
        collection, sorted in *idx* order.
        """
        placeholders =\
            [_Placeholder(sp) for sp in self.__shapes if sp.is_placeholder]
        placeholders.sort(key=lambda ph: ph.idx)
        return tuple(placeholders)

    @property
    def title(self):
        """The title shape in collection or None if no title placeholder."""
        for shape in self.__shapes:
            if shape._is_title:
                return shape
        return None

    def add_picture(self, file, left, top, width=None, height=None):
        """
        Add picture shape displaying image in *file*, where *file* can be
        either a path to a file (a string) or a file-like object.
        """
        image, rel = self.__slide._add_image(file)

        id = self.__next_shape_id
        name = 'Picture %d' % (id-1)
        desc = image._desc
        rId = rel._rId
        width, height = image._scale(width, height)

        pic = CT_Picture.new_pic(id, name, desc, rId, left, top, width, height)

        self.__spTree.append(pic)
        picture = _Picture(pic)
        self.__shapes.append(picture)
        return picture

    def add_shape(self, autoshape_type_id, left, top, width, height):
        """
        Add auto shape of type specified by *autoshape_type_id* (like
        ``MSO.SHAPE_RECTANGLE``) and of specified size at specified position.
        """
        autoshape_type = _AutoShapeType(autoshape_type_id)
        id_ = self.__next_shape_id
        name = '%s %d' % (autoshape_type.basename, id_-1)

        sp = CT_Shape.new_autoshape_sp(id_, name, autoshape_type.prst,
                                       left, top, width, height)
        shape = _Shape(sp)

        self.__spTree.append(sp)
        self.__shapes.append(shape)
        return shape

    def add_table(self, rows, cols, left, top, width, height):
        """
        Add table shape with the specified number of *rows* and *cols* at the
        specified position with the specified size. *width* is evenly
        distributed between the *cols* columns of the new table. Likewise,
        *height* is evenly distributed between the *rows* rows created.
        """
        id = self.__next_shape_id
        name = 'Table %d' % (id-1)
        graphicFrame = CT_GraphicalObjectFrame.new_table(
            id, name, rows, cols, left, top, width, height)
        self.__spTree.append(graphicFrame)
        table = _Table(graphicFrame)
        self.__shapes.append(table)
        return table

    def add_textbox(self, left, top, width, height):
        """
        Add text box shape of specified size at specified position.
        """
        id_ = self.__next_shape_id
        name = 'TextBox %d' % (id_-1)

        sp = CT_Shape.new_textbox_sp(id_, name, left, top, width, height)
        shape = _Shape(sp)

        self.__spTree.append(sp)
        self.__shapes.append(shape)
        return shape

    def _clone_layout_placeholders(self, slidelayout):
        """
        Add placeholder shapes based on those in *slidelayout*. Z-order of
        placeholders is preserved. Latent placeholders (date, slide number,
        and footer) are not cloned.
        """
        latent_ph_types = (PH_TYPE_DT, PH_TYPE_SLDNUM, PH_TYPE_FTR)
        for sp in slidelayout.shapes:
            if not sp.is_placeholder:
                continue
            ph = _Placeholder(sp)
            if ph.type in latent_ph_types:
                continue
            self.__clone_layout_placeholder(ph)

    def __clone_layout_placeholder(self, layout_ph):
        """
        Add a new placeholder shape based on the slide layout placeholder
        *layout_ph*.
        """
        id_ = self.__next_shape_id
        ph_type = layout_ph.type
        orient = layout_ph.orient
        shapename = self.__next_ph_name(ph_type, id_, orient)
        sz = layout_ph.sz
        idx = layout_ph.idx

        sp = CT_Shape.new_placeholder_sp(id_, shapename, ph_type, orient,
                                         sz, idx)
        shape = _Shape(sp)

        self.__spTree.append(sp)
        self.__shapes.append(shape)
        return shape

    def __next_ph_name(self, ph_type, id, orient):
        """
        Next unique placeholder name for placeholder shape of type *ph_type*,
        with id number *id* and orientation *orient*. Usually will be standard
        placeholder root name suffixed with id-1, e.g.
        __next_ph_name(PH_TYPE_TBL, 4, 'horz') ==> 'Table Placeholder 3'. The
        number is incremented as necessary to make the name unique within the
        collection. If *orient* is ``'vert'``, the placeholder name is
        prefixed with ``'Vertical '``.
        """
        basename = slide_ph_basenames[ph_type]
        # prefix rootname with 'Vertical ' if orient is 'vert'
        if orient == PH_ORIENT_VERT:
            basename = 'Vertical %s' % basename
        # increment numpart as necessary to make name unique
        numpart = id - 1
        names = self.__spTree.xpath('//p:cNvPr/@name', namespaces=_nsmap)
        while True:
            name = '%s %d' % (basename, numpart)
            if name not in names:
                break
            numpart += 1
        return name

    @property
    def __next_shape_id(self):
        """
        Next available drawing object id number in collection, starting from 1
        and making use of any gaps in numbering. In practice, the minimum id
        is 2 because the spTree element is always assigned id="1".
        """
        cNvPrs = self.__spTree.xpath('//p:cNvPr', namespaces=_nsmap)
        ids = [int(cNvPr.get('id')) for cNvPr in cNvPrs]
        ids.sort()
        # first gap in sequence wins, or falls off the end as max(ids)+1
        next_id = 1
        for id in ids:
            if id > next_id:
                break
            next_id += 1
        return next_id


class _Placeholder(object):
    """
    Decorator (pattern) class for adding placeholder properties to a shape
    that contains a placeholder element, e.g. ``<p:ph>``.
    """
    def __new__(cls, shape):
        cls = type('PlaceholderDecorator', (_Placeholder, shape.__class__), {})
        return object.__new__(cls)

    def __init__(self, shape):
        self.__decorated = shape
        xpath = './*[1]/p:nvPr/p:ph'
        self.__ph = self._element.xpath(xpath, namespaces=_nsmap)[0]

    def __getattr__(self, name):
        """
        Called when *name* is not found in ``self`` or in class tree. In this
        case, delegate attribute lookup to decorated (it's probably in its
        instance namespace).
        """
        return getattr(self.__decorated, name)

    @property
    def type(self):
        """Placeholder type, e.g. PH_TYPE_CTRTITLE"""
        return self.__ph.get('type', PH_TYPE_OBJ)

    @property
    def orient(self):
        """Placeholder 'orient' attribute, e.g. PH_ORIENT_HORZ"""
        return self.__ph.get('orient', PH_ORIENT_HORZ)

    @property
    def sz(self):
        """Placeholder 'sz' attribute, e.g. PH_SZ_FULL"""
        return self.__ph.get('sz', PH_SZ_FULL)

    @property
    def idx(self):
        """Placeholder 'idx' attribute, e.g. '0'"""
        return int(self.__ph.get('idx', 0))


class _Picture(_BaseShape):
    """
    A picture shape, one that places an image on a slide. Corresponds to the
    ``<p:pic>`` element.
    """
    def __init__(self, pic):
        super(_Picture, self).__init__(pic)

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, unconditionally
        ``MSO.PICTURE`` in this case.
        """
        return MSO.PICTURE


class _Shape(_BaseShape):
    """
    A shape that can appear on a slide. Corresponds to the ``<p:sp>`` element
    that can appear in any of the slide-type parts (slide, slideLayout,
    slideMaster, notesPage, notesMaster, handoutMaster).
    """
    def __init__(self, shape_element):
        super(_Shape, self).__init__(shape_element)

    @property
    def auto_shape_type(self):
        """
        Unique integer identifying the type of this auto shape, like
        ``MSO.SHAPE_ROUNDED_RECTANGLE``. Raises |ValueError| if this shape is
        not an auto shape.
        """
        sp = self._element
        if not sp.is_autoshape:
            msg = "shape is not an auto shape"
            raise ValueError(msg)
        prst = sp.prst
        auto_shape_type_id = _AutoShapeType._lookup_id_by_prst(prst)
        return auto_shape_type_id

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, like
        ``MSO.TEXT_BOX``.
        """
        if self.is_placeholder:
            return MSO.PLACEHOLDER
        sp = self._element
        if sp.is_autoshape:
            return MSO.AUTO_SHAPE
        if sp.is_textbox:
            return MSO.TEXT_BOX
        msg = '_Shape instance of unrecognized shape type'
        raise NotImplementedError(msg)


# ============================================================================
# Table-related classes
# ============================================================================

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
    def margin_bottom(self):
        """
        Integer value of bottom margin of cell in English Metric Units (EMU).
        If assigned |None|, the default value is used, 0.1 inches for left
        and right margins and 0.05 inches for top and bottom.
        """
        return self.__tc.marB

    @margin_bottom.setter
    def margin_bottom(self, margin_bottom):
        if (not isinstance(margin_bottom, (int, long))
                and margin_bottom is not None):
            tmpl = "margin_bottom must be integer or None, got '%s'"
            raise ValueError(tmpl % margin_bottom)
        self.__tc.marB = margin_bottom

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


# ============================================================================
# Text-related classes
# ============================================================================

class _TextFrame(object):
    """
    The part of a shape that contains its text. Not all shapes have a text
    frame. Corresponds to the ``<p:txBody>`` element that can appear as a
    child element of ``<p:sp>``. Not intended to be constructed directly.
    """
    def __init__(self, txBody):
        super(_TextFrame, self).__init__()
        self.__txBody = txBody

    @property
    def paragraphs(self):
        """
        Immutable sequence of |_Paragraph| instances corresponding to the
        paragraphs in this text frame. A text frame always contains at least
        one paragraph.
        """
        return tuple([_Paragraph(p) for p in self.__txBody[qn('a:p')]])

    def _set_text(self, text):
        """Replace all text in text frame with single run containing *text*"""
        self.clear()
        self.paragraphs[0].text = _to_unicode(text)

    #: Write-only. Assignment to *text* replaces all text currently contained
    #: in the text frame with the assigned expression. After assignment, the
    #: text frame contains exactly one paragraph containing a single run
    #: containing all the text. The assigned value can be a 7-bit ASCII
    #: string, a UTF-8 encoded 8-bit string, or unicode. String values are
    #: converted to unicode assuming UTF-8 encoding.
    text = property(None, _set_text)

    def _set_vertical_anchor(self, value):
        """
        Set ``anchor`` attribute of ``<a:bodyPr>`` element
        """
        value_map = {MSO.ANCHOR_TOP: 't', MSO.ANCHOR_MIDDLE: 'ctr',
                     MSO.ANCHOR_BOTTOM: 'b'}
        bodyPr = _get_or_add(self.__txBody, 'a:bodyPr')
        bodyPr.set('anchor', value_map[value])

    #: Write-only. Assignment to *vertical_anchor* sets the vertical
    #: alignment of the text frame to top, middle, or bottom. Valid values are
    #: ``MSO.ANCHOR_TOP``, ``MSO.ANCHOR_MIDDLE``, or ``MSO.ANCHOR_BOTTOM``.
    #: The ``MSO`` name is imported from ``pptx.constants``.
    vertical_anchor = property(None, _set_vertical_anchor)

    def add_paragraph(self):
        """
        Return new |_Paragraph| instance appended to the sequence of
        paragraphs contained in this text frame.
        """
        # <a:p> elements are last in txBody, so can simply append new one
        p = _Element('a:p', _nsmap)
        self.__txBody.append(p)
        return _Paragraph(p)

    def clear(self):
        """
        Remove all paragraphs except one empty one.
        """
        p_list = self.__txBody.xpath('./a:p', namespaces=_nsmap)
        for p in p_list[1:]:
            self.__txBody.remove(p)
        p = self.paragraphs[0]
        p.clear()


class _Font(object):
    """
    Character properties object, prominent among those properties being font
    size, font name, bold, italic, etc. Corresponds to ``<a:rPr>`` child
    element of a run. Also appears as ``<a:defRPr>`` and ``<a:endParaRPr>`` in
    paragraph and ``<a:defRPr>`` in list style elements. Not intended to be
    constructed directly.
    """
    def __init__(self, rPr):
        super(_Font, self).__init__()
        self.__rPr = rPr

    @property
    def bold(self):
        """
        Get or set boolean bold value of |_Font|, e.g.
        ``paragraph.font.bold = True``.
        """
        b = self.__rPr.get('b')
        return True if b in ('true', '1') else False

    @bold.setter
    def bold(self, bool):
        if bool:
            self.__rPr.set('b', '1')
        elif 'b' in self.__rPr.attrib:
            del self.__rPr.attrib['b']

    def _set_size(self, centipoints):
        # handle float centipoints value gracefully
        centipoints = int(centipoints)
        self.__rPr.set('sz', str(centipoints))

    #: Set the font size. In PresentationML, font size is expressed in
    #: hundredths of a point (centipoints). The :class:`pptx.util.Pt` class
    #: allows convenient conversion to centipoints from float or integer point
    #: values, e.g. ``Pt(12.5)``. I'm pretty sure I just made up the word
    #: *centipoint*, but it seems apt :).
    size = property(None, _set_size)


class _Paragraph(object):
    """
    Paragraph object. Not intended to be constructed directly.
    """
    def __init__(self, p):
        super(_Paragraph, self).__init__()
        self.__p = p

    def _get_alignment(self):
        """
        Return alignment type of this paragraph, e.g. ``PP.ALIGN_CENTER``.
        Can return |None|, meaning the paragraph has no alignment setting and
        its effective value is inherited from a higher-level object.
        """
        algn = self.__p.get_algn()
        return ParagraphAlignment.from_text_align_type(algn)

    def _set_alignment(self, alignment):
        """
        Set alignment of this paragraph to *alignment*, a constant value like
        ``PP.ALIGN_CENTER``. If *alignment* is None, any alignment setting is
        cleared and its effective value is inherited from a higher-level
        object.
        """
        algn = ParagraphAlignment.to_text_align_type(alignment)
        self.__p.set_algn(algn)

    #: Horizontal alignment of this paragraph, represented by a constant
    #: value like ``PP.ALIGN_CENTER``. Can be |None|, meaning the paragraph
    #: has no alignment setting and its effective value is inherited from a
    #: higher-level object.
    alignment = property(_get_alignment, _set_alignment)

    @property
    def font(self):
        """
        |_Font| object containing default character properties for the runs in
        this paragraph. These character properties override default properties
        inherited from parent objects such as the text frame the paragraph is
        contained in and they may be overridden by character properties set at
        the run level.
        """
        # A _Font instance is created on first access if it doesn't exist.
        # This can cause "litter" <a:pPr> and <a:defRPr> elements to be
        # included in the XML if the _Font element is referred to but not
        # populated with values.
        if not hasattr(self.__p, 'pPr'):
            pPr = _Element('a:pPr', _nsmap)
            self.__p.insert(0, pPr)
        if not hasattr(self.__p.pPr, 'defRPr'):
            _SubElement(self.__p.pPr, 'a:defRPr')
        return _Font(self.__p.pPr.defRPr)

    def _get_level(self):
        """
        Return integer indentation level of this paragraph.
        """
        if not hasattr(self.__p, 'pPr'):
            return 0
        return int(self.__p.pPr.get('lvl', 0))

    def _set_level(self, level):
        """
        Set indentation level of this paragraph to *level*, an integer value
        between 0 and 8 inclusive.
        """
        if not isinstance(level, int) or level < 0 or level > 8:
            msg = "paragraph level must be integer between 0 and 8 inclusive"
            raise ValueError(msg)
        if not hasattr(self.__p, 'pPr'):
            pPr = _Element('a:pPr', _nsmap)
            self.__p.insert(0, pPr)
        self.__p.pPr.set('lvl', str(level))

    #: Read-write integer indentation level of this paragraph. Range is 0-8.
    #: 0 represents a top-level paragraph and is the default value. Indentation
    #: level is most commonly encountered in a bulleted list, as is found on a
    #: word bullet slide.
    level = property(_get_level, _set_level)

    @property
    def runs(self):
        """
        Immutable sequence of |_Run| instances corresponding to the runs in
        this paragraph.
        """
        xpath = './a:r'
        r_elms = self.__p.xpath(xpath, namespaces=_nsmap)
        runs = []
        for r in r_elms:
            runs.append(_Run(r))
        return tuple(runs)

    def _set_text(self, text):
        """Replace runs with single run containing *text*"""
        self.clear()
        r = self.add_run()
        r.text = _to_unicode(text)

    #: Write-only. Assignment to *text* replaces all text currently contained
    #: in the paragraph. After assignment, the paragraph containins exactly
    #: one run containing the text value of the assigned expression. The
    #: assigned value can be a 7-bit ASCII string, a UTF-8 encoded 8-bit
    #: string, or unicode. String values are converted to unicode assuming
    #: UTF-8 encoding.
    text = property(None, _set_text)

    def add_run(self):
        """Return a new run appended to the runs in this paragraph."""
        r = _Element('a:r', _nsmap)
        _SubElement(r, 'a:t')
        # work out where to insert it, ahead of a:endParaRPr if there is one
        endParaRPr = _child(self.__p, 'a:endParaRPr')
        if endParaRPr is not None:
            endParaRPr.addprevious(r)
        else:
            self.__p.append(r)
        return _Run(r)

    def clear(self):
        """Remove all runs from this paragraph."""
        # retain pPr if present
        pPr = _child(self.__p, 'a:pPr')
        self.__p.clear()
        if pPr is not None:
            self.__p.insert(0, pPr)


class _Run(object):
    """
    Text run object. Corresponds to ``<a:r>`` child element in a paragraph.
    """
    def __init__(self, r):
        super(_Run, self).__init__()
        self.__r = r

    @property
    def font(self):
        """
        |_Font| object containing run-level character properties for the text
        in this run. Character properties can and perhaps most often are
        inherited from parent objects such as the paragraph and slide layout
        the run is contained in. Only those specifically assigned at the run
        level are contained in the |_Font| object.
        """
        if not hasattr(self.__r, 'rPr'):
            self.__r.insert(0, _Element('a:rPr', _nsmap))
        return _Font(self.__r.rPr)

    @property
    def text(self):
        """
        Read/Write. Text contained in the run. A regular text run is required
        to contain exactly one ``<a:t>`` (text) element. Assignment to *text*
        replaces the text currently contained in the run. The assigned value
        can be a 7-bit ASCII string, a UTF-8 encoded 8-bit string, or unicode.
        String values are converted to unicode assuming UTF-8 encoding.
        """
        return self.__r.t.text

    @text.setter
    def text(self, str):
        """Set the text of this run to *str*."""
        self.__r.t._setText(_to_unicode(str))
