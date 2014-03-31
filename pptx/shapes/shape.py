# encoding: utf-8

"""
Base shape-related objects such as BaseShape.
"""

from __future__ import absolute_import, print_function

from pptx.text import TextFrame
from pptx.util import to_unicode


class BaseShape(object):
    """
    Base class for shape objects, including |Shape|, |Picture|, and
    |GraphicFrame|.
    """
    def __init__(self, shape_elm, parent):
        super(BaseShape, self).__init__()
        self._element = shape_elm
        self._parent = parent

    @property
    def element(self):
        """
        Reference to the lxml element for this shape, e.g. a CT_Shape
        instance.
        """
        return self._element

    @property
    def has_textframe(self):
        """
        True if this shape has a txBody element and can contain text.
        """
        return self._element.txBody is not None

    @property
    def id(self):
        """
        Id of this shape. Note that ids are constrained to positive integers.
        """
        return self._element.shape_id

    @property
    def is_placeholder(self):
        """
        True if this shape is a placeholder. A shape is a placeholder if it
        has a <p:ph> element.
        """
        return self._element.has_ph_elm

    @property
    def name(self):
        """
        Name of this shape, e.g. 'Picture 7'
        """
        return self._element.shape_name

    @property
    def part(self):
        """
        The package part containing this object, a _BaseSlide subclass in
        this case.
        """
        return self._parent.part

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, like ``MSO.CHART``.
        Must be implemented by subclasses.
        """
        # # This one returns |None| unconditionally to account for shapes
        # # that haven't been implemented yet, like group shape and chart.
        # # Once those are done this should raise |NotImplementedError|.
        # msg = 'shape_type property must be implemented by subclasses'
        # raise NotImplementedError(msg)
        return None

    def _set_text(self, text):
        """Replace all text in shape with single run containing *text*"""
        if not self.has_textframe:
            raise TypeError("cannot set text of shape with no text frame")
        self.textframe.text = to_unicode(text)

    #: Write-only. Assignment to *text* replaces all text currently contained
    #: by the shape, resulting in a text frame containing exactly one
    #: paragraph, itself containing a single run. The assigned value can be a
    #: 7-bit ASCII string, a UTF-8 encoded 8-bit string, or unicode. String
    #: values are converted to unicode assuming UTF-8 encoding.
    text = property(None, _set_text)

    @property
    def textframe(self):
        """
        |TextFrame| instance for this shape. Raises |ValueError| if shape has
        no text frame. Use :attr:`has_textframe` to check whether a shape has
        a text frame.
        """
        txBody = self._element.txBody
        if txBody is None:
            raise ValueError('shape has no text frame')
        return TextFrame(txBody, self)
