# encoding: utf-8

"""
Picture shape.
"""

from .base import BaseShape
from ..dml.line import LineFormat
from ..enum.shapes import MSO_SHAPE_TYPE
from ..util import lazyproperty


class Picture(BaseShape):
    """
    A picture shape, one that places an image on a slide. Corresponds to the
    ``<p:pic>`` element.
    """
    def __init__(self, pic, parent):
        super(Picture, self).__init__(pic, parent)
        self._pic = pic

    @property
    def crop_bottom(self):
        """
        A |float| representing the relative portion cropped from the bottom
        of this picture where 1.0 represents 100%. For example, 25% is
        represented by 0.25. Negative values are valid as are values greater
        than 1.0.
        """
        return self._element.srcRect_b

    @property
    def crop_left(self):
        """
        A |float| representing the relative portion cropped from the left
        side of this picture where 1.0 represents 100%.
        """
        return self._element.srcRect_l

    @property
    def crop_right(self):
        """
        A |float| representing the relative portion cropped from the right
        side of this picture where 1.0 represents 100%.
        """
        return self._element.srcRect_r

    @property
    def crop_top(self):
        """
        A |float| representing the relative portion cropped from the top of
        this picture where 1.0 represents 100%.
        """
        return self._element.srcRect_t

    def get_or_add_ln(self):
        """
        Return the ``<a:ln>`` element containing the line format properties
        XML for this picture.
        """
        return self._pic.get_or_add_ln()

    @property
    def image(self):
        """
        An |Image| object providing access to the properties and bytes of the
        image in this picture shape.
        """
        slide, rId = self.part, self._element.blip_rId
        return slide.get_image(rId)

    @lazyproperty
    def line(self):
        """
        An instance of |LineFormat|, providing access to the properties of
        the outline bordering this picture, such as its color and width.
        """
        return LineFormat(self)

    @property
    def ln(self):
        """
        The ``<a:ln>`` element containing the line format properties such as
        line color and width. |None| if no ``<a:ln>`` element is present.
        """
        return self._pic.ln

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, unconditionally
        ``MSO_SHAPE_TYPE.PICTURE`` in this case.
        """
        return MSO_SHAPE_TYPE.PICTURE
