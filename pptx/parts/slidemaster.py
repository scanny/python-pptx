# encoding: utf-8

"""
Objects related to the slide master part
"""

from __future__ import absolute_import

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.part import PartCollection
from pptx.parts.slides import BaseSlide
from pptx.util import lazyproperty


class SlideMaster(BaseSlide):
    """
    Slide master part. Corresponds to package files
    ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    """
    @property
    def slide_layouts(self):
        """
        Sequence of |SlideLayout| objects belonging to this slide master
        """
        return _SlideLayouts()

    @lazyproperty
    def slidelayouts(self):
        """
        Collection of slide layout objects belonging to this slide master.
        """
        slidelayouts = PartCollection()
        sl_rels = [
            r for r in self.rels.values() if r.reltype == RT.SLIDE_LAYOUT
        ]
        for sl_rel in sl_rels:
            slide_layout = sl_rel.target_part
            slidelayouts.add_part(slide_layout)
        return slidelayouts


class _SlideLayouts(object):
    """
    Collection of slide layouts belonging to an instance of |SlideMaster|,
    having list access semantics. Supports indexed access, len(), and
    iteration.
    """
