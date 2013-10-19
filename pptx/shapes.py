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

from numbers import Number

from pptx.constants import MSO
from pptx.oxml import qn, CT_GraphicalObjectFrame, CT_Picture, CT_Shape
from pptx.shape import _BaseShape
from pptx.spec import (
    autoshape_types, namespaces, slide_ph_basenames
)
from pptx.spec import (
    PH_ORIENT_HORZ, PH_ORIENT_VERT, PH_SZ_FULL, PH_TYPE_DT, PH_TYPE_FTR,
    PH_TYPE_OBJ, PH_TYPE_SLDNUM
)
from pptx.table import _Table
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

class _AdjustmentCollection(object):
    """
    Sequence of |_Adjustment| for an auto shape, each representing an
    available adjustment for a shape of its type. Supports ``len()`` and
    indexed access, e.g. ``shape.adjustments[1] = 0.15``.
    """
    def __init__(self, prstGeom):
        super(_AdjustmentCollection, self).__init__()
        self.__adjustments = self.__initialized_adjustments(prstGeom)
        self.__prstGeom = prstGeom

    def __getitem__(self, key):
        """Provides indexed access, (e.g. 'adjustments[9]')."""
        return self.__adjustments[key].effective_value

    def __setitem__(self, key, value):
        """
        Provides item assignment via an indexed expression, e.g.
        ``adjustments[9] = 999.9``. Causes all adjustment values in
        collection to be written to the XML.
        """
        self.__adjustments[key].effective_value = value
        self.__rewrite_guides()

    def __initialized_adjustments(self, prstGeom):
        """
        Return an initialized list of adjustment values based on the contents
        of *prstGeom*
        """
        if prstGeom is None:
            return []
        davs = _AutoShapeType.default_adjustment_values(prstGeom.prst)
        adjustments = [_Adjustment(name, def_val) for name, def_val in davs]
        self.__update_adjustments_with_actuals(adjustments, prstGeom.gd)
        return adjustments

    def __rewrite_guides(self):
        """
        Write ``<a:gd>`` elements to the XML, one for each adjustment value.
        Any existing guide elements are overwritten.
        """
        guides = [(adj.name, adj.val) for adj in self.__adjustments]
        self.__prstGeom.rewrite_guides(guides)

    @staticmethod
    def __update_adjustments_with_actuals(adjustments, guides):
        """
        Update |_Adjustment| instances in *adjustments* with actual values
        held in *guides*, a list of ``<a:gd>`` elements. Guides with a name
        that does not match an adjustment object are skipped.
        """
        adjustments_by_name = dict((adj.name, adj) for adj in adjustments)
        for gd in guides:
            name = gd.get('name')
            actual = int(gd.get('fmla')[4:])
            try:
                adjustment = adjustments_by_name[name]
            except KeyError:
                continue
            adjustment.actual = actual
        return

    @property
    def _adjustments(self):
        """
        Sequence containing direct references to the |_Adjustment| objects
        contained in collection.
        """
        return tuple(self.__adjustments)

    def __len__(self):
        """Implement built-in function len()"""
        return len(self.__adjustments)


class _Adjustment(object):
    """
    An adjustment value for an autoshape.

    An adjustment value corresponds to the position of an adjustment handle on
    an auto shape. Adjustment handles are the small yellow diamond-shaped
    handles that appear on certain auto shapes and allow the outline of the
    shape to be adjusted. For example, a rounded rectangle has an adjustment
    handle that allows the radius of its corner rounding to be adjusted.

    Values are |float| and generally range from 0.0 to 1.0, although the value
    can be negative or greater than 1.0 in certain circumstances.
    """
    def __init__(self, name, def_val, actual=None):
        super(_Adjustment, self).__init__()
        self.name = name
        self.def_val = def_val
        self.actual = actual

    @property
    def effective_value(self):
        """
        Read/write |float| representing normalized adjustment value for this
        adjustment. Actual values are a large-ish integer expressed in shape
        coordinates, nominally between 0 and 100,000. The effective value is
        normalized to a corresponding value nominally between 0.0 and 1.0.
        Intuitively this represents the proportion of the width or height of
        the shape at which the adjustment value is located from its starting
        point. For simple shapes such as a rounded rectangle, this intuitive
        correspondence holds. For more complicated shapes and at more extreme
        shape proportions (e.g. width is much greater than height), the value
        can become negative or greater than 1.0.
        """
        raw_value = self.actual if self.actual is not None else self.def_val
        return self.__normalize(raw_value)

    @effective_value.setter
    def effective_value(self, value):
        if not isinstance(value, Number):
            tmpl = "adjustment value must be numeric, got '%s'"
            raise ValueError(tmpl % value)
        self.actual = self.__denormalize(value)

    @staticmethod
    def __denormalize(value):
        """
        Return integer corresponding to normalized *raw_value* on unit basis
        of 100,000. See _Adjustment.normalize for additional details.
        """
        return int(value * 100000.0)

    @staticmethod
    def __normalize(raw_value):
        """
        Return normalized value for *raw_value*. A normalized value is a
        |float| between 0.0 and 1.0 for nominal raw values between 0 and
        100,000. Raw values less than 0 and greater than 100,000 are valid
        and return values calculated on the same unit basis of 100,000.
        """
        return raw_value / 100000.0

    @property
    def val(self):
        """
        Denormalized effective value (expressed in shape coordinates),
        suitable for using in the XML.
        """
        return self.actual if self.actual is not None else self.def_val


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

    @staticmethod
    def default_adjustment_values(prst):
        """
        Return sequence of name, value tuples representing the adjustment
        value defaults for the auto shape type identified by *prst*.
        """
        try:
            autoshape_type_id = _AutoShapeType._lookup_id_by_prst(prst)
        except KeyError:
            return ()
        return autoshape_types[autoshape_type_id]['avLst']

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
    def __init__(self, sp):
        super(_Shape, self).__init__(sp)
        self.__adjustments = _AdjustmentCollection(sp.prstGeom)

    @property
    def adjustments(self):
        """
        Read-only reference to _AdjustmentsCollection instance for this
        shape
        """
        return self.__adjustments

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
