# encoding: utf-8

"""Autoshape-related objects such as Shape and Adjustment."""

from __future__ import absolute_import, division, print_function, unicode_literals

from numbers import Number
import xml.sax.saxutils as saxutils

from pptx.dml.fill import FillFormat
from pptx.dml.line import LineFormat
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE
from pptx.shapes.base import BaseShape
from pptx.spec import autoshape_types
from pptx.text.text import TextFrame
from pptx.util import lazyproperty


class Adjustment(object):
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
        super(Adjustment, self).__init__()
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
        raw_value = self.actual
        if raw_value is None:
            raw_value = self.def_val
        return self._normalize(raw_value)

    @effective_value.setter
    def effective_value(self, value):
        if not isinstance(value, Number):
            tmpl = "adjustment value must be numeric, got '%s'"
            raise ValueError(tmpl % value)
        self.actual = self._denormalize(value)

    @staticmethod
    def _denormalize(value):
        """
        Return integer corresponding to normalized *raw_value* on unit basis
        of 100,000. See Adjustment.normalize for additional details.
        """
        return int(value * 100000.0)

    @staticmethod
    def _normalize(raw_value):
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


class AdjustmentCollection(object):
    """
    Sequence of |Adjustment| instances for an auto shape, each representing
    an available adjustment for a shape of its type. Supports ``len()`` and
    indexed access, e.g. ``shape.adjustments[1] = 0.15``.
    """

    def __init__(self, prstGeom):
        super(AdjustmentCollection, self).__init__()
        self._adjustments_ = self._initialized_adjustments(prstGeom)
        self._prstGeom = prstGeom

    def __getitem__(self, key):
        """Provides indexed access, (e.g. 'adjustments[9]')."""
        return self._adjustments_[key].effective_value

    def __setitem__(self, key, value):
        """
        Provides item assignment via an indexed expression, e.g.
        ``adjustments[9] = 999.9``. Causes all adjustment values in
        collection to be written to the XML.
        """
        self._adjustments_[key].effective_value = value
        self._rewrite_guides()

    def _initialized_adjustments(self, prstGeom):
        """
        Return an initialized list of adjustment values based on the contents
        of *prstGeom*
        """
        if prstGeom is None:
            return []
        davs = AutoShapeType.default_adjustment_values(prstGeom.prst)
        adjustments = [Adjustment(name, def_val) for name, def_val in davs]
        self._update_adjustments_with_actuals(adjustments, prstGeom.gd_lst)
        return adjustments

    def _rewrite_guides(self):
        """
        Write ``<a:gd>`` elements to the XML, one for each adjustment value.
        Any existing guide elements are overwritten.
        """
        guides = [(adj.name, adj.val) for adj in self._adjustments_]
        self._prstGeom.rewrite_guides(guides)

    @staticmethod
    def _update_adjustments_with_actuals(adjustments, guides):
        """
        Update |Adjustment| instances in *adjustments* with actual values
        held in *guides*, a list of ``<a:gd>`` elements. Guides with a name
        that does not match an adjustment object are skipped.
        """
        adjustments_by_name = dict((adj.name, adj) for adj in adjustments)
        for gd in guides:
            name = gd.name
            actual = int(gd.fmla[4:])
            try:
                adjustment = adjustments_by_name[name]
            except KeyError:
                continue
            adjustment.actual = actual
        return

    @property
    def _adjustments(self):
        """
        Sequence containing direct references to the |Adjustment| objects
        contained in collection.
        """
        return tuple(self._adjustments_)

    def __len__(self):
        """Implement built-in function len()"""
        return len(self._adjustments_)


class AutoShapeType(object):
    """
    Return an instance of |AutoShapeType| containing metadata for an auto
    shape of type identified by *autoshape_type_id*. Instances are cached, so
    no more than one instance for a particular auto shape type is in memory.

    Instances provide the following attributes:

    .. attribute:: autoshape_type_id

       Integer uniquely identifying this auto shape type. Corresponds to a
       value in ``pptx.constants.MSO`` like ``MSO_SHAPE.ROUNDED_RECTANGLE``.

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

    _instances = {}

    def __new__(cls, autoshape_type_id):
        """
        Only create new instance on first call for content_type. After that,
        use cached instance.
        """
        # if there's not a matching instance in the cache, create one
        if autoshape_type_id not in cls._instances:
            inst = super(AutoShapeType, cls).__new__(cls)
            cls._instances[autoshape_type_id] = inst
        # return the instance; note that __init__() gets called either way
        return cls._instances[autoshape_type_id]

    def __init__(self, autoshape_type_id):
        """Initialize attributes from constant values in pptx.spec"""
        # skip loading if this instance is from the cache
        if hasattr(self, "_loaded"):
            return
        # raise on bad autoshape_type_id
        if autoshape_type_id not in autoshape_types:
            raise KeyError(
                "no autoshape type with id '%s' in pptx.spec.autoshape_types"
                % autoshape_type_id
            )
        # otherwise initialize new instance
        autoshape_type = autoshape_types[autoshape_type_id]
        self._autoshape_type_id = autoshape_type_id
        self._basename = autoshape_type["basename"]
        self._loaded = True

    @property
    def autoshape_type_id(self):
        """
        MSO_AUTO_SHAPE_TYPE enumeration value for this auto shape type
        """
        return self._autoshape_type_id

    @property
    def basename(self):
        """Base of shape name for this auto shape type.

        A shape name is like "Rounded Rectangle 7" and appears as an XML attribute for
        example at `p:sp/p:nvSpPr/p:cNvPr{name}`. This basename value is the name less
        the distinguishing integer. This value is escaped because at least one
        autoshape-type name includes double quotes ('"No" Symbol').
        """
        return saxutils.escape(self._basename, {'"': "&quot;"})

    @classmethod
    def default_adjustment_values(cls, prst):
        """
        Return sequence of name, value tuples representing the adjustment
        value defaults for the auto shape type identified by *prst*.
        """
        return autoshape_types[prst]["avLst"]

    @property
    def desc(self):
        """Informal description of this auto shape type"""
        return self._desc

    @classmethod
    def id_from_prst(cls, prst):
        """
        Return auto shape id (e.g. ``MSO_SHAPE.RECTANGLE``) corresponding to
        preset geometry keyword *prst*.
        """
        return MSO_AUTO_SHAPE_TYPE.from_xml(prst)

    @property
    def prst(self):
        """
        Preset geometry identifier string for this auto shape. Used in the
        ``prst`` attribute of ``<a:prstGeom>`` element to specify the geometry
        to be used in rendering the shape, for example ``'roundRect'``.
        """
        return MSO_AUTO_SHAPE_TYPE.to_xml(self._autoshape_type_id)


class Shape(BaseShape):
    """A shape that can appear on a slide.

    Corresponds to the ``<p:sp>`` element that can appear in any of the slide-type parts
    (slide, slideLayout, slideMaster, notesPage, notesMaster, handoutMaster).
    """

    def __init__(self, sp, parent):
        super(Shape, self).__init__(sp, parent)
        self._sp = sp

    @lazyproperty
    def adjustments(self):
        """
        Read-only reference to |AdjustmentCollection| instance for this
        shape
        """
        return AdjustmentCollection(self._sp.prstGeom)

    @property
    def auto_shape_type(self):
        """
        Enumeration value identifying the type of this auto shape, like
        ``MSO_SHAPE.ROUNDED_RECTANGLE``. Raises |ValueError| if this shape is
        not an auto shape.
        """
        if not self._sp.is_autoshape:
            raise ValueError("shape is not an auto shape")
        return self._sp.prst

    @lazyproperty
    def fill(self):
        """
        |FillFormat| instance for this shape, providing access to fill
        properties such as fill color.
        """
        return FillFormat.from_fill_parent(self._sp.spPr)

    def get_or_add_ln(self):
        """
        Return the ``<a:ln>`` element containing the line format properties
        XML for this shape.
        """
        return self._sp.get_or_add_ln()

    @property
    def has_text_frame(self):
        """
        |True| if this shape can contain text. Always |True| for an
        AutoShape.
        """
        return True

    @lazyproperty
    def line(self):
        """
        |LineFormat| instance for this shape, providing access to line
        properties such as line color.
        """
        return LineFormat(self)

    @property
    def ln(self):
        """
        The ``<a:ln>`` element containing the line format properties such as
        line color and width. |None| if no ``<a:ln>`` element is present.
        """
        return self._sp.ln

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, like
        ``MSO_SHAPE_TYPE.TEXT_BOX``.
        """
        if self.is_placeholder:
            return MSO_SHAPE_TYPE.PLACEHOLDER
        if self._sp.has_custom_geometry:
            return MSO_SHAPE_TYPE.FREEFORM
        if self._sp.is_autoshape:
            return MSO_SHAPE_TYPE.AUTO_SHAPE
        if self._sp.is_textbox:
            return MSO_SHAPE_TYPE.TEXT_BOX
        msg = "Shape instance of unrecognized shape type"
        raise NotImplementedError(msg)

    @property
    def text(self):
        """Read/write. Unicode (str in Python 3) representation of shape text.

        The returned string will contain a newline character (``"\\n"``) separating each
        paragraph and a vertical-tab (``"\\v"``) character for each line break (soft
        carriage return) in the shape's text.

        Assignment to *text* replaces all text previously contained in the shape, along
        with any paragraph or font formatting applied to it. A newline character
        (``"\\n"``) in the assigned text causes a new paragraph to be started.
        A vertical-tab (``"\\v"``) character in the assigned text causes a line-break
        (soft carriage-return) to be inserted. (The vertical-tab character appears in
        clipboard text copied from PowerPoint as its encoding of line-breaks.)

        Either bytes (Python 2 str) or unicode (Python 3 str) can be assigned. Bytes can
        be 7-bit ASCII or UTF-8 encoded 8-bit bytes. Bytes values are converted to
        unicode assuming UTF-8 encoding (which also works for ASCII).
        """
        return self.text_frame.text

    @text.setter
    def text(self, text):
        self.text_frame.text = text

    @property
    def text_frame(self):
        """|TextFrame| instance for this shape.

        Contains the text of the shape and provides access to text formatting
        properties.
        """
        txBody = self._element.get_or_add_txBody()
        return TextFrame(txBody, self)
