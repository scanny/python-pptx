# encoding: utf-8

"""
Slide layout-related objects.
"""

from __future__ import absolute_import

from ..opc.constants import RELATIONSHIP_TYPE as RT
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
