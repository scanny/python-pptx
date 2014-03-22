# encoding: utf-8

"""
Slide layout-related objects.
"""

from __future__ import absolute_import

from warnings import warn

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.parts.slide import BaseSlide
from pptx.shapes.placeholder import BasePlaceholder, BasePlaceholders
from pptx.shapes.shapetree import BaseShapeFactory, BaseShapeTree
from pptx.util import lazyproperty


class SlideLayout(BaseSlide):
    """
    Slide layout part. Corresponds to package files
    ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """
    @lazyproperty
    def placeholders(self):
        """
        Instance of |_LayoutPlaceholders| containing sequence of placeholder
        shapes in this slide layout, sorted in *idx* order.
        """
        return _LayoutPlaceholders(self)

    @lazyproperty
    def shapes(self):
        """
        Instance of |_LayoutShapeTree| containing sequence of shapes
        appearing on this slide layout.
        """
        return _LayoutShapeTree(self)

    @property
    def slide_master(self):
        """
        Slide master from which this slide layout inherits properties.
        """
        return self.part_related_by(RT.SLIDE_MASTER)

    @property
    def slidemaster(self):
        """
        Deprecated. Use ``.slide_master`` property instead.
        """
        msg = (
            'SlideLayout.slidemaster property is deprecated. Use .slide_mast'
            'er instead.'
        )
        warn(msg, UserWarning, stacklevel=2)
        return self.slide_master


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
        Return an instance of the appropriate shape proxy class for
        *shape_elm* on a slide layout.
        """
        tag_name = shape_elm.tag
        if tag_name == qn('p:sp') and shape_elm.has_ph_elm:
            return _LayoutPlaceholder(shape_elm, parent)
        return BaseShapeFactory(shape_elm, parent)


class _LayoutPlaceholder(BasePlaceholder):
    """
    Placeholder shape on a slide layout, providing differentiated behavior
    for slide layout placeholders, in particular, inheriting shape properties
    from the master placeholder having the same type.
    """


class _LayoutPlaceholders(BasePlaceholders):
    """
    Sequence of _LayoutPlaceholder instances representing the placeholder
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
