# encoding: utf-8

"""
Presentation part, the main part in a .pptx package.
"""

from __future__ import absolute_import

from ..opc.package import XmlPart
from ..opc.packuri import PackURI
from ..presentation import Presentation
from ..util import lazyproperty


class PresentationPart(XmlPart):
    """
    Top level class in object model, represents the contents of the /ppt
    directory of a .pptx file.
    """
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
