# encoding: utf-8

"""
Picture shape.
"""

from pptx.constants import MSO
from pptx.shapes.shape import BaseShape


class Picture(BaseShape):
    """
    A picture shape, one that places an image on a slide. Corresponds to the
    ``<p:pic>`` element.
    """
    def __init__(self, pic, parent):
        super(Picture, self).__init__(pic, parent)

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, unconditionally
        ``MSO.PICTURE`` in this case.
        """
        return MSO.PICTURE
