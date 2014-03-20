# encoding: utf-8

"""
Presentation part, the main part in a .pptx package.
"""

from __future__ import absolute_import

from warnings import warn

from pptx.opc.package import Part
from pptx.oxml import parse_xml_bytes
from pptx.parts.slide import SlideCollection
from pptx.util import lazyproperty


class PresentationPart(Part):
    """
    Top level class in object model, represents the contents of the /ppt
    directory of a .pptx file.
    """
    def __init__(self, partname, content_type, presentation_elm, package):
        super(PresentationPart, self).__init__(
            partname, content_type, element=presentation_elm, package=package
        )

    @classmethod
    def load(cls, partname, content_type, blob, package):
        presentation_elm = parse_xml_bytes(blob)
        presentation = cls(partname, content_type, presentation_elm, package)
        return presentation

    @property
    def sldMasterIdLst(self):
        """
        The ``<p:sldMasterIdLst>`` child element specifying the slide masters
        of this presentation in the XML.
        """
        return self._element.get_or_add_sldMasterIdLst()

    @lazyproperty
    def slide_masters(self):
        """
        Sequence of |SlideMaster| objects belonging to this presentation
        """
        return _SlideMasters(self)

    @property
    def slidemasters(self):
        """
        Deprecated. Use ``.slide_masters`` property instead.
        """
        msg = (
            'Presentation.slidemasters property is deprecated. Use .slide_ma'
            'sters instead.'
        )
        warn(msg, UserWarning, stacklevel=2)
        return self.slide_masters

    @lazyproperty
    def slides(self):
        """
        |SlideCollection| object containing the slides in this presentation.
        """
        sldIdLst = self._element.get_or_add_sldIdLst()
        slides = SlideCollection(sldIdLst, self)
        slides.rename_slides()  # start from known state
        return slides


class _SlideMasters(object):
    """
    Collection of |SlideMaster| instances belonging to a presentation. Has
    list access semantics, supporting indexed access, len(), and iteration.
    """
    def __init__(self, presentation):
        super(_SlideMasters, self).__init__()
        self._presentation = presentation

    def __getitem__(self, idx):
        """
        Provide indexed access, (e.g. ``slide_masters[2]``).
        """
        sldMasterId_lst = self._sldMasterIdLst.sldMasterId_lst
        if idx >= len(sldMasterId_lst):
            raise IndexError('slide master index out of range')
        rId = sldMasterId_lst[idx].rId
        return self._presentation.related_parts[rId]

    def __iter__(self):
        """
        Generate a reference to each of the |SlideMaster| instances in the
        collection, in sequence.
        """
        for rId in self._iter_rIds():
            yield self._presentation.related_parts[rId]

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slide_masters) == 4').
        """
        return len(self._sldMasterIdLst)

    def _iter_rIds(self):
        """
        Generate the rId for each slide master in the collection, in
        sequence.
        """
        sldMasterId_lst = self._sldMasterIdLst.sldMasterId_lst
        for sldMasterId in sldMasterId_lst:
            yield sldMasterId.rId

    @property
    def _sldMasterIdLst(self):
        """
        The ``<p:sldMasterIdLst>`` element specifying the slide masters in
        this collection. This element is a child of the ``<p:presentation>``
        element, the root element of a presentation part.
        """
        return self._presentation.sldMasterIdLst
