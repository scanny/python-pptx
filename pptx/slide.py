# encoding: utf-8

"""
Slide-related objects.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .opc.constants import RELATIONSHIP_TYPE as RT
from .opc.packuri import PackURI
from .shapes.factory import SlidePlaceholders
from .shapes.shapetree import SlideShapeTree
from .shared import ParentedElementProxy
from .util import lazyproperty


class Slide(ParentedElementProxy):
    """
    Slide object. Provides access to shapes and slide-level properties.
    """
    @property
    def name(self):
        """
        Internal name of this slide.
        """
        return self._element.cSld.name

    @lazyproperty
    def placeholders(self):
        """
        Instance of |SlidePlaceholders| containing sequence of placeholder
        shapes in this slide.
        """
        return SlidePlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self):
        """
        Instance of |SlideShapeTree| containing sequence of shape objects
        appearing on this slide.
        """
        return SlideShapeTree(self)

    @property
    def slide_layout(self):
        """
        |SlideLayout| object this slide inherits appearance from.
        """
        return self.part.slide_layout

    @property
    def spTree(self):
        """
        Reference to ``<p:spTree>`` element for this slide
        """
        return self._element.cSld.spTree


from .parts.slide import SlidePart


class Slides(ParentedElementProxy):
    """
    Sequence of slides belonging to an instance of |Presentation|, having
    list semantics for access to individual slides. Supports indexed access,
    len(), and iteration.
    """
    def __init__(self, sldIdLst, prs):
        super(Slides, self).__init__(sldIdLst, prs)
        self._sldIdLst = sldIdLst

    def __getitem__(self, idx):
        """
        Provide indexed access, (e.g. 'slides[0]').
        """
        if idx >= len(self._sldIdLst):
            raise IndexError('slide index out of range')
        sldId = self._sldIdLst[idx]
        slide_part = self.part.related_parts[sldId.rId]
        return slide_part.slide

    def __iter__(self):
        """
        Support iteration (e.g. 'for slide in slides:').
        """
        for sldId in self._sldIdLst:
            slide_part = self.part.related_parts[sldId.rId]
            yield slide_part.slide

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slides) == 4').
        """
        return len(self._sldIdLst)

    def add_slide(self, slide_layout):
        """
        Return a newly added slide that inherits layout from *slide_layout*.
        """
        # TODO: Refactor me
        partname = self._next_partname
        package = self.part.package
        slide_part = SlidePart.new(slide_layout, partname, package)
        rId = self.part.relate_to(slide_part, RT.SLIDE)
        self._sldIdLst.add_sldId(rId)
        return slide_part.slide

    def index(self, slide):
        """
        Map *slide* to an integer representing its zero-based position in
        this slide collection. Raises |ValueError| on *slide* not present.
        """
        sld = slide._element
        for idx, this_slide in enumerate(self):
            if this_slide._element is sld:
                return idx
        raise ValueError('%s is not in slide collection' % slide)

    def rename_slides(self):
        """
        Assign partnames like ``/ppt/slides/slide9.xml`` to all slides in the
        collection. The name portion is always ``slide``. The number part
        forms a continuous sequence starting at 1 (e.g. 1, 2, 3, ...). The
        extension is always ``.xml``.
        """
        for idx, slide in enumerate(self):
            slide.partname = PackURI('/ppt/slides/slide%d.xml' % (idx+1))

    @property
    def _next_partname(self):
        """
        Return |PackURI| instance containing the partname for a slide to be
        appended to this slide collection, e.g. ``/ppt/slides/slide9.xml``
        for a slide collection containing 8 slides.
        """
        partname_str = '/ppt/slides/slide%d.xml' % (len(self)+1)
        return PackURI(partname_str)


class SlideLayouts(ParentedElementProxy):
    """
    Collection of slide layouts belonging to an instance of |SlideMaster|,
    having list access semantics. Supports indexed access, len(), and
    iteration.
    """

    __slots__ = ('_sldLayoutIdLst',)

    def __init__(self, sldLayoutIdLst, parent):
        super(SlideLayouts, self).__init__(sldLayoutIdLst, parent)
        self._sldLayoutIdLst = sldLayoutIdLst

    def __getitem__(self, idx):
        """
        Provide indexed access, (e.g. ``slide_layouts[2]``).
        """
        try:
            sldLayoutId = self._sldLayoutIdLst[idx]
        except IndexError:
            raise IndexError('slide layout index out of range')
        return self.part.related_parts[sldLayoutId.rId]

    def __iter__(self):
        """
        Generate a reference to each of the |SlideLayout| instances in the
        collection, in sequence.
        """
        for sldLayoutId in self._sldLayoutIdLst:
            yield self.part.related_parts[sldLayoutId.rId]

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slides) == 4').
        """
        return len(self._sldLayoutIdLst)


class SlideMasters(ParentedElementProxy):
    """
    Collection of |SlideMaster| instances belonging to a presentation. Has
    list access semantics, supporting indexed access, len(), and iteration.
    """

    __slots__ = ('_sldMasterIdLst',)

    def __init__(self, sldMasterIdLst, parent):
        super(SlideMasters, self).__init__(sldMasterIdLst, parent)
        self._sldMasterIdLst = sldMasterIdLst

    def __getitem__(self, idx):
        """
        Provide indexed access, (e.g. ``slide_masters[2]``).
        """
        try:
            sldMasterId = self._sldMasterIdLst[idx]
        except IndexError:
            raise IndexError('slide master index out of range')
        return self.part.related_parts[sldMasterId.rId]

    def __iter__(self):
        """
        Generate a reference to each of the |SlideMaster| instances in the
        collection, in sequence.
        """
        for smi in self._sldMasterIdLst:
            yield self.part.related_parts[smi.rId]

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slide_masters) == 4').
        """
        return len(self._sldMasterIdLst)
