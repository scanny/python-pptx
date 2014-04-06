# encoding: utf-8

"""
Picture shape.
"""

from ..enum.shapes import MSO_SHAPE_TYPE
from .shape import BaseShape


class Picture(BaseShape):
    """
    A picture shape, one that places an image on a slide. Corresponds to the
    ``<p:pic>`` element.
    """
    def __init__(self, pic, parent):
        super(Picture, self).__init__(pic, parent)
        self._pic = pic

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, unconditionally
        ``MSO_SHAPE_TYPE.PICTURE`` in this case.
        """
        return MSO_SHAPE_TYPE.PICTURE
