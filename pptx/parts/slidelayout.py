# encoding: utf-8

"""
Slide objects, including Slide and SlideMaster.
"""

from __future__ import absolute_import

from warnings import warn

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.slide import BaseSlide
from pptx.shapes.placeholder import BasePlaceholder
from pptx.shapes.shapetree import BaseShapeTree


class SlideLayout(BaseSlide):
    """
    Slide layout part. Corresponds to package files
    ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """
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


class _LayoutPlaceholder(BasePlaceholder):
    """
    Placeholder shape on a slide layout, providing differentiated behavior
    for slide layout placeholders, in particular, inheriting shape properties
    from the master placeholder having the same type.
    """
