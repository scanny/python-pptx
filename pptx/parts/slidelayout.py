# encoding: utf-8

"""
Slide layout-related objects.
"""

from __future__ import absolute_import

from warnings import warn

from ..enum.shapes import PP_PLACEHOLDER
from ..opc.constants import RELATIONSHIP_TYPE as RT
from ..oxml.ns import qn
from ..shapes.placeholder import LayoutPlaceholder
from ..shapes.shapetree import (
    BasePlaceholders, BaseShapeFactory, BaseShapeTree
)
from .slide import BaseSlide
from ..util import lazyproperty


class SlideLayout(BaseSlide):
    """
    Slide layout part. Corresponds to package files
    ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """
    def iter_cloneable_placeholders(self):
        """
        Generate a reference to each layout placeholder on this slide layout
        that should be cloned to a slide when the layout is applied to the
        slide.
        """
        latent_ph_types = (
            PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.SLIDE_NUMBER
        )
        for ph in self.placeholders:
            if ph.ph_type not in latent_ph_types:
                yield ph

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
