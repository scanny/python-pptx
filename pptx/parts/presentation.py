# encoding: utf-8

"""
Presentation part, the main part in a .pptx package.
"""

from __future__ import absolute_import

from ..opc.constants import RELATIONSHIP_TYPE as RT
from ..opc.package import XmlPart
from ..opc.packuri import PackURI
from ..presentation import Presentation
from .slide import SlidePart
from ..util import lazyproperty


class PresentationPart(XmlPart):
    """
    Top level class in object model, represents the contents of the /ppt
    directory of a .pptx file.
    """
    def add_slide(self, slide_layout):
        """
        Return an (rId, slide) pair of a newly created blank slide that
        inherits appearance from *slide_layout*.
        """
        partname = self._next_slide_partname
        slide_layout_part = slide_layout.part
        slide_part = SlidePart.new(partname, self.package, slide_layout_part)
        rId = self.relate_to(slide_part, RT.SLIDE)
        return rId, slide_part.slide

    @property
    def core_properties(self):
        """
        A |CoreProperties| object providing read/write access to the core
        properties of this presentation.
        """
        return self.package.core_properties

    @lazyproperty
    def presentation(self):
        """
        A |Presentation| object providing access to the content of this
        presentation.
        """
        return Presentation(self._element, self)

    def related_slide(self, rId):
        """
        Return the |Slide| object for the related |SlidePart| corresponding
        to relationship key *rId*.
        """
        return self.related_parts[rId].slide

    def related_slide_master(self, rId):
        """
        Return the |SlideMaster| object for the related |SlideMasterPart|
        corresponding to relationship key *rId*.
        """
        return self.related_parts[rId].slide_master

    def rename_slide_parts(self, rIds):
        """
        Assign incrementing partnames like ``/ppt/slides/slide9.xml`` to the
        slide parts identified by *rIds*, in the order their id appears in
        that sequence. The name portion is always ``slide``. The number part
        forms a continuous sequence starting at 1 (e.g. 1, 2, ... 10, ...).
        The extension is always ``.xml``.
        """
        for idx, rId in enumerate(rIds):
            slide_part = self.related_parts[rId]
            slide_part.partname = PackURI(
                '/ppt/slides/slide%d.xml' % (idx+1)
            )

    def save(self, path_or_stream):
        """
        Save this presentation package to *path_or_stream*, which can be
        either a path to a filesystem location (a string) or a file-like
        object.
        """
        self.package.save(path_or_stream)

    @property
    def _next_slide_partname(self):
        """
        Return |PackURI| instance containing the partname for a slide to be
        appended to this slide collection, e.g. ``/ppt/slides/slide9.xml``
        for a slide collection containing 8 slides.
        """
        sldIdLst = self._element.get_or_add_sldIdLst()
        partname_str = '/ppt/slides/slide%d.xml' % (len(sldIdLst)+1)
        return PackURI(partname_str)
