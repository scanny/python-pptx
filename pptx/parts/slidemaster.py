# encoding: utf-8

"""
Objects related to the slide master part
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..shapes.shapetree import MasterPlaceholders, MasterShapes
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
        Instance of |MasterPlaceholders| containing sequence of placeholder
        shapes in this slide master, sorted in *idx* order.
        """
        return MasterPlaceholders(self._element.spTree, self)

    def related_slide_layout(self, rId):
        """
        Return the |SlideLayout| object of the related |SlideLayoutPart|
        corresponding to relationship key *rId*.
        """
        return self.related_parts[rId].slide_layout

    @lazyproperty
    def shapes(self):
        """
        Instance of |MasterShapes| containing sequence of shape objects
        appearing on this slide.
        """
        return MasterShapes(self._element.spTree, self)

    @lazyproperty
    def slide_layouts(self):
        """
        Sequence of |SlideLayout| objects belonging to this slide master
        """
        return SlideLayouts(self._element.get_or_add_sldLayoutIdLst(), self)
