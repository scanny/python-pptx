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
        self._pic = pic

    @property
    def height(self):
        """
        Distance between top and bottom extents of shape in integer EMUs.
        """
        return self._ext.cy

    @height.setter
    def height(self, value):
        self._ext.cy = value

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, unconditionally
        ``MSO.PICTURE`` in this case.
        """
        return MSO.PICTURE

    @property
    def width(self):
        """
        Distance between left and right extents of shape in integer EMUs.
        """
        return self._ext.cx

    @width.setter
    def width(self, value):
        self._ext.cx = value

    @property
    def _ext(self):
        """
        Get or add sp.spPr.xfrm.ext element
        """
        spPr = self._pic.spPr
        xfrm = spPr.get_or_add_xfrm()
        return xfrm.get_or_add_ext()
