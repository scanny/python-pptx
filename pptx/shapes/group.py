# encoding: utf-8

"""
Group shape and related objects
"""

from .base import BaseShape
from ..enum.shapes import MSO_SHAPE_TYPE
from ..util import lazyproperty

class GroupShape(BaseShape):
    """
    Group shape.

    Corresponds to a ``<p:grpSp>`` element in the shape tree.
    """
    __slots__ = ('_shapes',)

    def __init__(self, shape_elm, parent):
        super(GroupShape, self).__init__(shape_elm, parent)
        self.x_min = self.y_min = self.x_max = self.y_max = None

    def _min(self, values):
        return int(min([x for x in values if x is not None]))

    def _max(self, values):
        return int(max([x for x in values if x is not None]))

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, e.g.
        ``MSO_SHAPE_TYPE.TABLE``.
        """
        return MSO_SHAPE_TYPE.GROUP

    @lazyproperty
    def shapes(self):
        """
        Instance of |GroupShapes| containing sequence of shape objects
        appearing on this group.
        """
        # Local import to avoid circular references
        from .shapetree import GroupShapes
        return GroupShapes(self._element, self)

    def _adjustExtents(self, left, top, width, height):
        self.x_min = self._min([self.x_min, left])
        self.y_min = self._min([self.y_min, top])
        self.x_max = self._max([self.x_max, left + width])
        self.y_max = self._max([self.y_max, top + height])

        self._element.add_or_replace_extents(self.x_min, self.y_min, self.x_max-self.x_min, self.y_max-self.y_min)
