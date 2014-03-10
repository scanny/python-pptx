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
        return _SlideLayouts(self)

    @property
    def sldLayoutIdLst(self):
        """
        The ``<p:sldLayoutIdLst>`` chile element specifying the slide layouts
        of this slide master in the XML.
        """
        raise NotImplementedError

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
    def __init__(self, slide_master):
        super(_SlideLayouts, self).__init__()
        self._slide_master = slide_master

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slides) == 4').
        """
        return len(self._sldLayoutIdLst)

    @property
    def _sldLayoutIdLst(self):
        """
        The ``<p:sldLayoutIdLst>`` element specifying the slide layouts in
        this collection. This element is a child of the ``<p:sldMaster>``
        element, the root element of a slide master part.
        """
        return self._slide_master.sldLayoutIdLst
