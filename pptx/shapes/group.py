# encoding: utf-8

"""
Group shape.
"""

from __future__ import absolute_import, print_function

from ..util import lazyproperty
from .base import BaseShape


class Group(BaseShape):
    """
    A group shape. Contains an arbitrary number of other shape objects.
    """
    def __init__(self, el, parent):
        super(Group, self).__init__(el, parent)

    @lazyproperty
    def shapes(self):
        from .shapetree import GroupShapeTree
        return GroupShapeTree(self, slide=self._parent._slide)
