# encoding: utf-8

"""
Slide objects, including Slide and SlideMaster.
"""

from __future__ import absolute_import

from warnings import warn

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.slide import BaseSlide


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
