# encoding: utf-8

"""
Objects related to the slide master part
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..oxml.ns import qn
from ..shapes.placeholder import MasterPlaceholder
from ..shapes.shapetree import (
    BasePlaceholders, BaseShapeFactory, BaseShapeTree
)
from .slide import BaseSlidePart
from ..slide import SlideLayouts
from ..util import lazyproperty


class SlideMasterPart(BaseSlidePart):
    """
    Slide master part. Corresponds to package files
    ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    """
    @lazyproperty
    def placeholders(self):
        """
        Instance of |_MasterPlaceholders| containing sequence of placeholder
        shapes in this slide master, sorted in *idx* order.
        """
        return _MasterPlaceholders(self._element.spTree, self)

    def related_slide_layout(self, rId):
        """
        Return the |SlideLayout| object of the related |SlideLayoutPart|
        corresponding to relationship key *rId*.
        """
        return self.related_parts[rId].slide_layout

    @lazyproperty
    def shapes(self):
        """
        Instance of |_MasterShapeTree| containing sequence of shape objects
        appearing on this slide.
        """
        return _MasterShapeTree(self._element.spTree, self)

    @lazyproperty
    def slide_layouts(self):
        """
        Sequence of |SlideLayout| objects belonging to this slide master
        """
        return SlideLayouts(self._element.get_or_add_sldLayoutIdLst(), self)


class _MasterShapeTree(BaseShapeTree):
    """
    Sequence of shapes appearing on a slide master. The first shape in the
    sequence is the backmost in z-order and the last shape is topmost.
    Supports indexed access, len(), and iteration.
    """
    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        return _MasterShapeFactory(shape_elm, self)


def _MasterShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*
    on a slide master.
    """
    tag_name = shape_elm.tag
    if tag_name == qn('p:sp') and shape_elm.has_ph_elm:
        return MasterPlaceholder(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


class _MasterPlaceholders(BasePlaceholders):
    """
    Sequence of _MasterPlaceholder instances representing the placeholder
    shapes on a slide master.
    """
    def get(self, ph_type, default=None):
        """
        Return the first placeholder shape with type *ph_type* (e.g. 'body'),
        or *default* if no such placeholder shape is present in the
        collection.
        """
        for placeholder in self:
            if placeholder.ph_type == ph_type:
                return placeholder
        return default

    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        return _MasterShapeFactory(shape_elm, self)
