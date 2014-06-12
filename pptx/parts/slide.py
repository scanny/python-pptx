# encoding: utf-8

"""
Slide and related objects.
"""

from __future__ import absolute_import

from warnings import warn

from ..enum.shapes import PP_PLACEHOLDER
from ..opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from ..opc.package import Part
from ..opc.packuri import PackURI
from ..oxml import parse_xml
from ..oxml.ns import _nsmap, qn
from ..oxml.simpletypes import ST_Direction
from ..oxml.slide import CT_Slide
from ..shapes.autoshape import AutoShapeType
from ..shapes.placeholder import BasePlaceholder, BasePlaceholders
from ..shapes.shapetree import BaseShapeFactory, BaseShapeTree
from ..util import lazyproperty


class BaseSlide(Part):
    """
    Base class for slide parts, e.g. slide, slideLayout, slideMaster,
    notesSlide, notesMaster, and handoutMaster.
    """
    def __init__(self, partname, content_type, element, package):
        super(BaseSlide, self).__init__(
            partname, content_type, element=element, package=package
        )

    @classmethod
    def load(cls, partname, content_type, blob, package):
        slide_elm = parse_xml(blob)
        slide = cls(partname, content_type, slide_elm, package)
        return slide

    @property
    def name(self):
        """
        Internal name of this slide.
        """
        return self._element.cSld.name

    @property
    def part(self):
        """
        Part of the parent protocol, "children" of the slide will not know
        the part that contains them so must ask their parent object. That
        chain of delegation ends here for slide child objects.
        """
        return self

    @property
    def spTree(self):
        """
        Reference to ``<p:spTree>`` element for this slide
        """
        spTree_lst = self._element.xpath(
            './p:cSld/p:spTree', namespaces=_nsmap
        )
        return spTree_lst[0]

    def _add_image(self, img_file):
        """
        Return 2-tuple ``(image, rId)`` representing an |Image| part
        corresponding to the image in *img_file*, newly created if no
        matching image part is already present. If the slide already has a
        relationship to an existing image, that relationship is reused.
        """
        image = self._package._images.add_image(img_file)
        rId = self.relate_to(image, RT.IMAGE)
        return (image, rId)


class Slide(BaseSlide):
    """
    Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml.
    """
    @classmethod
    def new(cls, slidelayout, partname, package):
        """
        Return a new slide based on *slidelayout* and having *partname*,
        created from scratch.
        """
        slide_elm = CT_Slide.new()
        slide = cls(partname, CT.PML_SLIDE, slide_elm, package)
        slide.shapes.clone_layout_placeholders(slidelayout)
        slide.relate_to(slidelayout, RT.SLIDE_LAYOUT)
        return slide

    @lazyproperty
    def placeholders(self):
        """
        Instance of |_SlidePlaceholders| containing sequence of placeholder
        shapes in this slide.
        """
        return _SlidePlaceholders(self)

    @lazyproperty
    def shapes(self):
        """
        Instance of |_SlideShapeTree| containing sequence of shape objects
        appearing on this slide.
        """
        return _SlideShapeTree(self)

    @property
    def slide_layout(self):
        """
        |SlideLayout| object this slide inherits appearance from.
        """
        return self.part_related_by(RT.SLIDE_LAYOUT)

    @property
    def slidelayout(self):
        """
        Deprecated. Use ``.slide_layout`` property instead.
        """
        msg = (
            'Slide.slidelayout property is deprecated. Use .slide_layout '
            'instead.'
        )
        warn(msg, UserWarning, stacklevel=2)
        return self.slide_layout


class SlideCollection(object):
    """
    Sequence of slides belonging to an instance of |Presentation|, having list
    semantics for access to individual slides. Supports indexed access,
    len(), and iteration.
    """
    def __init__(self, sldIdLst, prs):
        super(SlideCollection, self).__init__()
        self._sldIdLst = sldIdLst
        self._prs = prs

    def __getitem__(self, idx):
        """
        Provide indexed access, (e.g. 'slides[0]').
        """
        if idx >= len(self._sldIdLst):
            raise IndexError('slide index out of range')
        rId = self._sldIdLst[idx].rId
        return self._prs.related_parts[rId]

    def __iter__(self):
        """
        Support iteration (e.g. 'for slide in slides:').
        """
        for sldId in self._sldIdLst:
            rId = sldId.rId
            yield self._prs.related_parts[rId]

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slides) == 4').
        """
        return len(self._sldIdLst)

    def add_slide(self, slidelayout):
        """
        Return a newly added slide that inherits layout from *slidelayout*.
        """
        partname = self._next_partname
        package = self._prs.package
        slide = Slide.new(slidelayout, partname, package)
        rId = self._prs.relate_to(slide, RT.SLIDE)
        self._sldIdLst.add_sldId(rId)
        return slide

    def rename_slides(self):
        """
        Assign partnames like ``/ppt/slides/slide9.xml`` to all slides in the
        collection. The name portion is always ``slide``. The number part
        forms a continuous sequence starting at 1 (e.g. 1, 2, 3, ...). The
        extension is always ``.xml``.
        """
        for idx, slide in enumerate(self):
            partname_str = '/ppt/slides/slide%d.xml' % (idx+1)
            slide.partname = PackURI(partname_str)

    @property
    def _next_partname(self):
        """
        Return |PackURI| instance containing the partname for a slide to be
        appended to this slide collection, e.g. ``/ppt/slides/slide9.xml``
        for a slide collection containing 8 slides.
        """
        partname_str = '/ppt/slides/slide%d.xml' % (len(self)+1)
        return PackURI(partname_str)


class _SlideShapeTree(BaseShapeTree):
    """
    Sequence of shapes appearing on a slide. The first shape in the sequence
    is the backmost in z-order and the last shape is topmost. Supports indexed
    access, len(), index(), and iteration.
    """
    def add_picture(self, image_file, left, top, width=None, height=None):
        """
        Add picture shape displaying image in *image_file*, where
        *image_file* can be either a path to a file (a string) or a file-like
        object.
        """
        image_part, rId = self._get_or_add_image_part(image_file)
        pic = self._add_pic_from_image_part(
            image_part, rId, left, top, width, height
        )
        picture = self._shape_factory(pic)
        return picture

    def add_shape(self, autoshape_type_id, left, top, width, height):
        """
        Add auto shape of type specified by *autoshape_type_id* (like
        ``MSO_SHAPE.RECTANGLE``) and of specified size at specified position.
        """
        autoshape_type = AutoShapeType(autoshape_type_id)
        sp = self._add_sp_from_autoshape_type(
            autoshape_type, left, top, width, height
        )
        shape = self._shape_factory(sp)
        return shape

    def add_table(self, rows, cols, left, top, width, height):
        """
        Add table shape with the specified number of *rows* and *cols* at the
        specified position with the specified size. *width* is evenly
        distributed between the *cols* columns of the new table. Likewise,
        *height* is evenly distributed between the *rows* rows created.
        """
        graphicFrame = self._add_graphicFrame_containing_table(
            rows, cols, left, top, width, height
        )
        table = self._shape_factory(graphicFrame)
        return table

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
        Instance of |_SlidePlaceholders| containing sequence of placeholder
        shapes in this slide.
        """
        return self._slide.placeholders

    @property
    def title(self):
        """
        The title placeholder shape on the slide or |None| if the slide has
        no title placeholder.
        """
        for elm in self._spTree.iter_shape_elms():
            if elm.ph_idx == 0:
                return self._shape_factory(elm)
        return None

    def _add_graphicFrame_containing_table(self, rows, cols, x, y, cx, cy):
        """
        Return a newly added ``<p:graphicFrame>`` element containing a table
        as specified by the parameters.
        """
        id_ = self._next_shape_id
        name = 'Table %d' % (id_-1)
        graphicFrame = self._spTree.add_table(
            id_, name, rows, cols, x, y, cx, cy
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
        desc = image_part._desc
        scaled_cx, scaled_cy = image_part._scale(cx, cy)

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

    def _get_or_add_image_part(self, image_file):
        """
        Return an (image_part, rId) 2-tuple corresponding to an image part
        containing the image in *image_file*, and related to this object's
        part with the key *rId*. If the image part and/or relationship
        already exists, they are reused, otherwise they are newly created.
        """
        slide = self._slide
        image_part, rId = slide._add_image(image_file)
        return image_part, rId

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
        names = self._spTree.xpath('//p:cNvPr/@name', namespaces=_nsmap)
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
        return _SlideShapeFactory(shape_elm, self)


def _SlideShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*
    on a slide.
    """
    tag_name = shape_elm.tag
    if tag_name == qn('p:sp') and shape_elm.has_ph_elm:
        return _SlidePlaceholder(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


class _SlidePlaceholder(BasePlaceholder):
    """
    Placeholder shape on a slide. Inherits shape properties from its
    corresponding slide layout placeholder.
    """
    @property
    def height(self):
        """
        The effective height of this placeholder shape; its directly-applied
        height if it has one, otherwise the height of its parent layout
        placeholder.
        """
        return self._effective_value('height')

    @property
    def left(self):
        """
        The effective left of this placeholder shape; its directly-applied
        left if it has one, otherwise the left of its parent layout
        placeholder.
        """
        return self._effective_value('left')

    @property
    def top(self):
        """
        The effective top of this placeholder shape; its directly-applied
        top if it has one, otherwise the top of its parent layout
        placeholder.
        """
        return self._effective_value('top')

    @property
    def width(self):
        """
        The effective width of this placeholder shape; its directly-applied
        width if it has one, otherwise the width of its parent layout
        placeholder.
        """
        return self._effective_value('width')

    def _effective_value(self, attr_name):
        """
        The effective value of *attr_name* on this placeholder shape; its
        directly-applied value if it has one, otherwise the value on the
        layout placeholder it inherits from.
        """
        directly_applied_value = getattr(
            super(_SlidePlaceholder, self), attr_name
        )
        if directly_applied_value is not None:
            return directly_applied_value
        return self._inherited_value(attr_name)

    def _inherited_value(self, attr_name):
        """
        The attribute value, e.g. 'width' of the layout placeholder this
        slide placeholder inherits from
        """
        layout_placeholder = self._layout_placeholder
        if layout_placeholder is None:
            return None
        inherited_value = getattr(layout_placeholder, attr_name)
        return inherited_value

    @property
    def _layout_placeholder(self):
        """
        The layout placeholder shape this slide placeholder inherits from
        """
        layout = self._slide_layout
        layout_placeholder = layout.placeholders.get(idx=self.idx)
        return layout_placeholder

    @property
    def _slide_layout(self):
        """
        The slide layout from which the slide this placeholder belongs to
        inherits.
        """
        slide = self.part
        return slide.slide_layout


class _SlidePlaceholders(BasePlaceholders):
    """
    Sequence of placeholder shapes on a slide.
    """
    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm* on a slide.
        """
        return _SlideShapeFactory(shape_elm, self)
