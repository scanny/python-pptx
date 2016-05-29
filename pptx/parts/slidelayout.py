# encoding: utf-8

"""
Slide layout-related objects.
"""

from __future__ import absolute_import

from ..opc.constants import RELATIONSHIP_TYPE as RT
from ..oxml.ns import qn
from ..shapes.placeholder import LayoutPlaceholder
from ..shapes.shapetree import (
    BasePlaceholders, BaseShapeFactory, BaseShapeTree
)
from .slide import BaseSlidePart
from ..slide import SlideLayout
from ..util import lazyproperty


class SlideLayoutPart(BaseSlidePart):
    """
    Slide layout part. Corresponds to package files
    ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """
    @lazyproperty
    def slide_layout(self):
        """
        The |SlideLayout| object representing this part.
        """
        return SlideLayout(self._element, self)

    @property
    def slide_master(self):
        """
        Slide master from which this slide layout inherits properties.
        """
        return self.part_related_by(RT.SLIDE_MASTER)


class _LayoutShapeTree(BaseShapeTree):
    """
    Sequence of shapes appearing on a slide layout. The first shape in the
    sequence is the backmost in z-order and the last shape is topmost.
    Supports indexed access, len(), index(), and iteration.
    """
    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        parent = self
        return _LayoutShapeFactory(shape_elm, parent)


def _LayoutShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*
    on a slide layout.
    """
    tag_name = shape_elm.tag
    if tag_name == qn('p:sp') and shape_elm.has_ph_elm:
        return LayoutPlaceholder(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


class _LayoutPlaceholders(BasePlaceholders):
    """
    Sequence of |LayoutPlaceholder| instances representing the placeholder
    shapes on a slide layout.
    """
    def get(self, idx, default=None):
        """
        Return the first placeholder shape with matching *idx* value, or
        *default* if not found.
        """
        for placeholder in self:
            if placeholder.idx == idx:
                return placeholder
        return default

    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        parent = self
        return _LayoutShapeFactory(shape_elm, parent)
