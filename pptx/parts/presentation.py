# encoding: utf-8

"""
Presentation part, the main part in a .pptx package.
"""

from __future__ import absolute_import

from warnings import warn

from ..opc.constants import RELATIONSHIP_TYPE as RT
from ..opc.package import XmlPart
from ..opc.packuri import PackURI
from .slide import Slide
from ..util import lazyproperty


class PresentationPart(XmlPart):
    """
    Top level class in object model, represents the contents of the /ppt
    directory of a .pptx file.
    """
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

    @property
    def slide_height(self):
        """
        Height of slides in this presentation, in English Metric Units (EMU)
        """
        sldSz = self._element.sldSz
        return sldSz.cy

    @slide_height.setter
    def slide_height(self, height):
        sldSz = self._element.sldSz
        sldSz.cy = height

    @property
    def slide_width(self):
        """
        Width of slides in this presentation, in English Metric Units (EMU)
        """
        sldSz = self._element.sldSz
        return sldSz.cx

    @slide_width.setter
    def slide_width(self, width):
        sldSz = self._element.sldSz
        sldSz.cx = width

    @lazyproperty
    def slides(self):
        """
        |_Slides| object containing the slides in this presentation.
        """
        sldIdLst = self._element.get_or_add_sldIdLst()
        slides = _Slides(sldIdLst, self)
        slides.rename_slides()  # start from known state
        return slides


class _Slides(object):
    """
    Sequence of slides belonging to an instance of |Presentation|, having list
    semantics for access to individual slides. Supports indexed access,
    len(), and iteration.
    """
    def __init__(self, sldIdLst, prs):
        super(_Slides, self).__init__()
        self._sldIdLst = sldIdLst
        self._prs = prs

    def __getitem__(self, idx):
        """
        Provide indexed access, (e.g. 'slides[0]').
        """
        if idx >= len(self._sldIdLst):
            raise IndexError('slide index out of range')
        rId = self._sldIdLst[idx].rId
        return self._prs.related_parts[rId]

    def __iter__(self):
        """
        Support iteration (e.g. 'for slide in slides:').
        """
        for sldId in self._sldIdLst:
            rId = sldId.rId
            yield self._prs.related_parts[rId]

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slides) == 4').
        """
        return len(self._sldIdLst)

    def add_slide(self, slidelayout):
        """
        Return a newly added slide that inherits layout from *slidelayout*.
        """
        partname = self._next_partname
        package = self._prs.package
        slide = Slide.new(slidelayout, partname, package)
        rId = self._prs.relate_to(slide, RT.SLIDE)
        self._sldIdLst.add_sldId(rId)
        return slide

    def index(self, item):
        """
        Map *item* to an integer representing its zero-based position in this
        slide collection. Raises |ValueError| if *item* is not a slide in
        this collection.
        """
        item_id = id(item)
        for idx, slide in enumerate(self):
            if id(slide) == item_id:
                return idx
        raise ValueError('%s is not in slide collection' % item)

    def rename_slides(self):
        """
        Assign partnames like ``/ppt/slides/slide9.xml`` to all slides in the
        collection. The name portion is always ``slide``. The number part
        forms a continuous sequence starting at 1 (e.g. 1, 2, 3, ...). The
        extension is always ``.xml``.
        """
        for idx, slide in enumerate(self):
            partname_str = '/ppt/slides/slide%d.xml' % (idx+1)
            slide.partname = PackURI(partname_str)

    @property
    def _next_partname(self):
        """
        Return |PackURI| instance containing the partname for a slide to be
        appended to this slide collection, e.g. ``/ppt/slides/slide9.xml``
        for a slide collection containing 8 slides.
        """
        partname_str = '/ppt/slides/slide%d.xml' % (len(self)+1)
        return PackURI(partname_str)


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
