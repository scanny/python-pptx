# encoding: utf-8

"""
The shape tree, the structure that holds a slide's shapes.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .autoshape import AutoShapeType, Shape
from .base import BaseShape
from .connector import Connector
from ..enum.shapes import PP_PLACEHOLDER
from .graphfrm import GraphicFrame
from ..oxml.ns import qn
from ..oxml.shapes.graphfrm import CT_GraphicalObjectFrame
from ..oxml.simpletypes import ST_Direction
from .picture import Picture
from .placeholder import (
    ChartPlaceholder, LayoutPlaceholder, MasterPlaceholder,
    PicturePlaceholder, PlaceholderGraphicFrame, PlaceholderPicture,
    SlidePlaceholder, TablePlaceholder
)
from ..shared import ParentedElementProxy


def BaseShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*.
    """
    shape_cls = {
        qn('p:cxnSp'):        Connector,
        qn('p:sp'):           Shape,
        qn('p:pic'):          Picture,
        qn('p:graphicFrame'): GraphicFrame,
    }.get(shape_elm.tag, BaseShape)

    return shape_cls(shape_elm, parent)


class _BaseShapes(ParentedElementProxy):
    """
    Base class for a shape collection appearing in a slide-type object,
    include Slide, SlideLayout, and SlideMaster, providing common methods.
    """
    def __init__(self, spTree, parent):
        super(_BaseShapes, self).__init__(spTree, parent)
        self._spTree = spTree

    def __getitem__(self, idx):
        """
        Return shape at *idx* in sequence, e.g. ``shapes[2]``.
        """
        shape_elms = list(self._iter_member_elms())
        try:
            shape_elm = shape_elms[idx]
        except IndexError:
            raise IndexError('shape index out of range')
        return self._shape_factory(shape_elm)

    def __iter__(self):
        """
        Generate a reference to each shape in the collection, in sequence.
        """
        for shape_elm in self._iter_member_elms():
            yield self._shape_factory(shape_elm)

    def __len__(self):
        """
        Return count of shapes in this shape tree. A group shape contributes
        1 to the total, without regard to the number of shapes contained in
        the group.
        """
        shape_elms = list(self._iter_member_elms())
        return len(shape_elms)

    @staticmethod
    def _is_member_elm(shape_elm):
        """
        Return true if *shape_elm* represents a member of this collection,
        False otherwise.
        """
        return True

    def _iter_member_elms(self):
        """
        Generate each child of the ``<p:spTree>`` element that corresponds to
        a shape, in the sequence they appear in the XML.
        """
        for shape_elm in self._spTree.iter_shape_elms():
            if self._is_member_elm(shape_elm):
                yield shape_elm

    @property
    def _next_shape_id(self):
        """
        Next available positive integer drawing object id in shape tree,
        starting from 1 and making use of any gaps in numbering. In practice,
        the minimum id is 2 because the spTree element is always assigned
        id="1".
        """
        id_str_lst = self._spTree.xpath('//@id')
        used_ids = [int(id_str) for id_str in id_str_lst if id_str.isdigit()]
        for n in range(1, len(used_ids)+2):
            if n not in used_ids:
                return n

    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        return BaseShapeFactory(shape_elm, self)


class BasePlaceholders(_BaseShapes):
    """
    Base class for placeholder collections that differentiate behaviors for
    a master, layout, and slide.
    """
    @staticmethod
    def _is_member_elm(shape_elm):
        """
        True if *shape_elm* is a placeholder shape, False otherwise.
        """
        return shape_elm.has_ph_elm


class LayoutPlaceholders(BasePlaceholders):
    """
    Sequence of |LayoutPlaceholder| instances representing the placeholder
    shapes on a slide layout.
    """
    def get(self, idx, default=None):
        """
        Return the first placeholder shape with matching *idx* value, or
        *default* if not found.
        """
        for placeholder in self:
            if placeholder.idx == idx:
                return placeholder
        return default

    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        return _LayoutShapeFactory(shape_elm, self)


def _LayoutShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*
    on a slide layout.
    """
    tag_name = shape_elm.tag
    if tag_name == qn('p:sp') and shape_elm.has_ph_elm:
        return LayoutPlaceholder(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


class LayoutShapes(_BaseShapes):
    """
    Sequence of shapes appearing on a slide layout. The first shape in the
    sequence is the backmost in z-order and the last shape is topmost.
    Supports indexed access, len(), index(), and iteration.
    """
    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        return _LayoutShapeFactory(shape_elm, self)


class MasterPlaceholders(BasePlaceholders):
    """
    Sequence of _MasterPlaceholder instances representing the placeholder
    shapes on a slide master.
    """
    def get(self, ph_type, default=None):
        """
        Return the first placeholder shape with type *ph_type* (e.g. 'body'),
        or *default* if no such placeholder shape is present in the
        collection.
        """
        for placeholder in self:
            if placeholder.ph_type == ph_type:
                return placeholder
        return default

    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        return _MasterShapeFactory(shape_elm, self)


class MasterShapes(_BaseShapes):
    """
    Sequence of shapes appearing on a slide master. The first shape in the
    sequence is the backmost in z-order and the last shape is topmost.
    Supports indexed access, len(), and iteration.
    """
    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        return _MasterShapeFactory(shape_elm, self)


def _MasterShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*
    on a slide master.
    """
    tag_name = shape_elm.tag
    if tag_name == qn('p:sp') and shape_elm.has_ph_elm:
        return MasterPlaceholder(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


def _SlidePlaceholderFactory(shape_elm, parent):
    """
    Return a placeholder shape of the appropriate type for *shape_elm*.
    """
    tag = shape_elm.tag
    if tag == qn('p:sp'):
        Constructor = {
            PP_PLACEHOLDER.BITMAP:  PicturePlaceholder,
            PP_PLACEHOLDER.CHART:   ChartPlaceholder,
            PP_PLACEHOLDER.PICTURE: PicturePlaceholder,
            PP_PLACEHOLDER.TABLE:   TablePlaceholder,
        }.get(shape_elm.ph_type, SlidePlaceholder)
    elif tag == qn('p:graphicFrame'):
        Constructor = PlaceholderGraphicFrame
    elif tag == qn('p:pic'):
        Constructor = PlaceholderPicture
    else:
        Constructor = BaseShapeFactory
    return Constructor(shape_elm, parent)


class SlidePlaceholders(ParentedElementProxy):
    """
    Collection of placeholder shapes on a slide. Supports iteration,
    :func:`len`, and dictionary-style lookup on the `idx` value of the
    placeholders it contains.
    """

    __slots__ = ()

    def __getitem__(self, idx):
        """
        Access placeholder shape having *idx*. Note that while this looks
        like list access, idx is actually a dictionary key and will raise
        |KeyError| if no placeholder with that idx value is in the
        collection.
        """
        for e in self._element.iter_ph_elms():
            if e.ph_idx == idx:
                return SlideShapeFactory(e, self)
        raise KeyError('no placeholder on this slide with idx == %d' % idx)

    def __iter__(self):
        """
        Generate placeholder shapes in `idx` order.
        """
        ph_elms = sorted(
            [e for e in self._element.iter_ph_elms()], key=lambda e: e.ph_idx
        )
        return (SlideShapeFactory(e, self) for e in ph_elms)

    def __len__(self):
        """
        Return count of placeholder shapes.
        """
        return len(list(self._element.iter_ph_elms()))


def SlideShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*
    on a slide.
    """
    if shape_elm.has_ph_elm:
        return _SlidePlaceholderFactory(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


class SlideShapes(_BaseShapes):
    """
    Sequence of shapes appearing on a slide. The first shape in the sequence
    is the backmost in z-order and the last shape is topmost. Supports indexed
    access, len(), index(), and iteration.
    """
    def add_chart(self, chart_type, x, y, cx, cy, chart_data):
        """
        Add a new chart of *chart_type* to the slide, positioned at (*x*,
        *y*), having size (*cx*, *cy*), and depicting *chart_data*.
        *chart_type* is one of the :ref:`XlChartType` enumeration values.
        *chart_data* is a |ChartData| object populated with the categories
        and series values for the chart. Note that a |GraphicFrame| shape
        object is returned, not the |Chart| object contained in that graphic
        frame shape. The chart object may be accessed using the :attr:`chart`
        property of the returned |GraphicFrame| object.
        """
        rId = self.part.add_chart_part(chart_type, chart_data)
        graphic_frame = self._add_chart_graphic_frame(rId, x, y, cx, cy)
        return graphic_frame

    def add_connector(self, connector_type, begin_x, begin_y, end_x, end_y):
        """
        Add a newly created connector shape to the end of this shape tree.
        *connector_type* is a member of the :ref:`MsoConnectorType`
        enumeration and the end-point values are specified as EMU values. The
        returned connector is of type *connector_type* and has begin and end
        points as specified.
        """
        cxnSp = self._add_cxnSp(
            connector_type, begin_x, begin_y, end_x, end_y
        )
        return self._shape_factory(cxnSp)

    def add_picture(self, image_file, left, top, width=None, height=None):
        """
        Add picture shape displaying image in *image_file*, where
        *image_file* can be either a path to a file (a string) or a file-like
        object.
        """
        image_part, rId = self.part.get_or_add_image_part(image_file)
        pic = self._add_pic_from_image_part(
            image_part, rId, left, top, width, height
        )
        return self._shape_factory(pic)

    def add_shape(self, autoshape_type_id, left, top, width, height):
        """
        Add auto shape of type specified by *autoshape_type_id* (like
        ``MSO_SHAPE.RECTANGLE``) and of specified size at specified position.
        """
        autoshape_type = AutoShapeType(autoshape_type_id)
        sp = self._add_sp_from_autoshape_type(
            autoshape_type, left, top, width, height
        )
        return self._shape_factory(sp)

    def add_table(self, rows, cols, left, top, width, height):
        """
        Add a |GraphicFrame| object containing a table with the specified
        number of *rows* and *cols* and the specified position and size.
        *width* is evenly distributed between the columns of the new table.
        Likewise, *height* is evenly distributed between the rows. Note that
        the ``.table`` property on the returned |GraphicFrame| shape must be
        used to access the enclosed |Table| object.
        """
        graphicFrame = self._add_graphicFrame_containing_table(
            rows, cols, left, top, width, height
        )
        graphic_frame = self._shape_factory(graphicFrame)
        return graphic_frame

    def add_textbox(self, left, top, width, height):
        """
        Add text box shape of specified size at specified position on slide.
        """
        sp = self._add_textbox_sp(left, top, width, height)
        textbox = self._shape_factory(sp)
        return textbox

    def clone_layout_placeholders(self, slide_layout):
        """
        Add placeholder shapes based on those in *slide_layout*. Z-order of
        placeholders is preserved. Latent placeholders (date, slide number,
        and footer) are not cloned.
        """
        for placeholder in slide_layout.iter_cloneable_placeholders():
            self._clone_layout_placeholder(placeholder)

    def index(self, shape):
        """
        Return the index of *shape* in this sequence, raising |ValueError| if
        *shape* is not in the collection.
        """
        shape_elm = shape.element
        for idx, elm in enumerate(self._spTree.iter_shape_elms()):
            if elm is shape_elm:
                return idx
        raise ValueError('shape not in collection')

    @property
    def placeholders(self):
        """
        Instance of |SlidePlaceholders| containing sequence of placeholder
        shapes in this slide.
        """
        return self.parent.placeholders

    @property
    def title(self):
        """
        The title placeholder shape on the slide or |None| if the slide has
        no title placeholder.
        """
        for elm in self._spTree.iter_ph_elms():
            if elm.ph_idx == 0:
                return self._shape_factory(elm)
        return None

    def _add_chart_graphicFrame(self, rId, x, y, cx, cy):
        """
        Add a new ``<p:graphicFrame>`` element to this shape tree having the
        specified position and size and referring to the chart part
        identified by *rId*.
        """
        shape_id = self._next_shape_id
        name = 'Chart %d' % (shape_id-1)
        graphicFrame = CT_GraphicalObjectFrame.new_chart_graphicFrame(
            shape_id, name, rId, x, y, cx, cy
        )
        self._spTree.append(graphicFrame)
        return graphicFrame

    def _add_chart_graphic_frame(self, rId, x, y, cx, cy):
        """
        Return a |GraphicFrame| object having the specified position and size
        and referring to the chart part identified by *rId*.
        """
        graphicFrame = self._add_chart_graphicFrame(rId, x, y, cx, cy)
        graphic_frame = self._shape_factory(graphicFrame)
        return graphic_frame

    def _add_cxnSp(self, connector_type, begin_x, begin_y, end_x, end_y):
        """
        Return a newly-added `p:cxnSp` element for a connector of
        *connector_type* beginning at (*begin_x*, *begin_y*) and extending to
        (*end_x*, *end_y*).
        """
        id_ = self._next_shape_id
        name = 'Connector %d' % (id_-1)

        flipH, flipV = begin_x > end_x, begin_y > end_y
        x, y = min(begin_x, end_x), min(begin_y, end_y)
        cx, cy = abs(end_x - begin_x), abs(end_y - begin_y)

        return self._spTree.add_cxnSp(
            id_, name, connector_type, x, y, cx, cy, flipH, flipV
        )

    def _add_graphicFrame_containing_table(self, rows, cols, x, y, cx, cy):
        """
        Return a newly added ``<p:graphicFrame>`` element containing a table
        as specified by the parameters.
        """
        _id = self._next_shape_id
        name = 'Table %d' % (_id-1)
        graphicFrame = self._spTree.add_table(
            _id, name, rows, cols, x, y, cx, cy
        )
        return graphicFrame

    def _add_pic_from_image_part(self, image_part, rId, x, y, cx, cy):
        """
        Return a newly added ``<p:pic>`` element specifying a picture shape
        displaying *image_part* with size and position specified by *x*, *y*,
        *cx*, and *cy*. The element is appended to the shape tree, causing it
        to be displayed first in z-order on the slide.
        """
        id = self._next_shape_id
        name = 'Picture %d' % (id-1)
        desc = image_part.desc
        scaled_cx, scaled_cy = image_part.scale(cx, cy)

        pic = self._spTree.add_pic(
            id, name, desc, rId, x, y, scaled_cx, scaled_cy
        )

        return pic

    def _add_sp_from_autoshape_type(self, autoshape_type, x, y, cx, cy):
        """
        Return a newly-added ``<p:sp>`` element for a shape of
        *autoshape_type* at position (x, y) and of size (cx, cy).
        """
        id_ = self._next_shape_id
        name = '%s %d' % (autoshape_type.basename, id_-1)
        sp = self._spTree.add_autoshape(
            id_, name, autoshape_type.prst, x, y, cx, cy
        )
        return sp

    def _add_textbox_sp(self, x, y, cx, cy):
        """
        Return a newly-added textbox ``<p:sp>`` element at position (x, y)
        and of size (cx, cy).
        """
        id_ = self._next_shape_id
        name = 'TextBox %d' % (id_-1)
        sp = self._spTree.add_textbox(id_, name, x, y, cx, cy)
        return sp

    def _clone_layout_placeholder(self, layout_placeholder):
        """
        Add a new placeholder shape based on the slide layout placeholder
        *layout_ph*.
        """
        id_ = self._next_shape_id
        ph_type = layout_placeholder.ph_type
        orient = layout_placeholder.orient
        name = self._next_ph_name(ph_type, id_, orient)
        sz = layout_placeholder.sz
        idx = layout_placeholder.idx

        self._spTree.add_placeholder(id_, name, ph_type, orient, sz, idx)

    def _next_ph_name(self, ph_type, id, orient):
        """
        Next unique placeholder name for placeholder shape of type *ph_type*,
        with id number *id* and orientation *orient*. Usually will be standard
        placeholder root name suffixed with id-1, e.g.
        _next_ph_name(ST_PlaceholderType.TBL, 4, 'horz') ==>
        'Table Placeholder 3'. The number is incremented as necessary to make
        the name unique within the collection. If *orient* is ``'vert'``, the
        placeholder name is prefixed with ``'Vertical '``.
        """
        basename = {
            # BODY is named 'Notes Placeholder' in a notes master
            PP_PLACEHOLDER.BODY:         'Text Placeholder',
            PP_PLACEHOLDER.CHART:        'Chart Placeholder',
            PP_PLACEHOLDER.BITMAP:       'ClipArt Placeholder',
            PP_PLACEHOLDER.CENTER_TITLE: 'Title',
            PP_PLACEHOLDER.ORG_CHART:    'SmartArt Placeholder',
            PP_PLACEHOLDER.DATE:         'Date Placeholder',
            PP_PLACEHOLDER.FOOTER:       'Footer Placeholder',
            PP_PLACEHOLDER.HEADER:       'Header Placeholder',
            PP_PLACEHOLDER.MEDIA_CLIP:   'Media Placeholder',
            PP_PLACEHOLDER.OBJECT:       'Content Placeholder',
            PP_PLACEHOLDER.PICTURE:      'Picture Placeholder',
            # PP_PLACEHOLDER.SLIDE_IMAGE:  'Slide Image Placeholder',
            PP_PLACEHOLDER.SLIDE_NUMBER: 'Slide Number Placeholder',
            PP_PLACEHOLDER.SUBTITLE:     'Subtitle',
            PP_PLACEHOLDER.TABLE:        'Table Placeholder',
            PP_PLACEHOLDER.TITLE:        'Title',
        }[ph_type]

        # prefix rootname with 'Vertical ' if orient is 'vert'
        if orient == ST_Direction.VERT:
            basename = 'Vertical %s' % basename

        # increment numpart as necessary to make name unique
        numpart = id - 1
        names = self._spTree.xpath('//p:cNvPr/@name')
        while True:
            name = '%s %d' % (basename, numpart)
            if name not in names:
                break
            numpart += 1

        return name

    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        return SlideShapeFactory(shape_elm, self)
