# encoding: utf-8

"""
Slide-related objects.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .enum.shapes import PP_PLACEHOLDER
from .shapes.shapetree import (
    LayoutPlaceholders, LayoutShapes, MasterPlaceholders, MasterShapes,
    SlidePlaceholders, SlideShapes
)
from .shared import ParentedElementProxy, PartElementProxy
from .util import lazyproperty


class _BaseSlide(PartElementProxy):
    """
    Slide object. Provides access to shapes and slide-level properties.
    """

    __slots__ = ()

    @property
    def name(self):
        """
        String representing the internal name of this slide. Returns an empty
        string (`''`) if no name is assigned. Assigning an empty string or
        |None| to this property causes any name to be removed.
        """
        return self._element.cSld.name

    @name.setter
    def name(self, value):
        new_value = '' if value is None else value
        self._element.cSld.name = new_value


class Slide(_BaseSlide):
    """
    Slide object. Provides access to shapes and slide-level properties.
    """

    __slots__ = ('_placeholders', '_shapes')

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
        Instance of |SlideShapes| containing sequence of shape objects
        appearing on this slide.
        """
        return SlideShapes(self._element.spTree, self)

    @property
    def slide_id(self):
        """
        The integer value that uniquely identifies this slide within this
        presentation. The slide id does not change if the position of this
        slide in the slide sequence is changed by adding, rearranging, or
        deleting slides.
        """
        return self.part.slide_id

    @property
    def slide_layout(self):
        """
        |SlideLayout| object this slide inherits appearance from.
        """
        return self.part.slide_layout


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
        try:
            sldId = self._sldIdLst[idx]
        except IndexError:
            raise IndexError('slide index out of range')
        return self.part.related_slide(sldId.rId)

    def __iter__(self):
        """
        Support iteration (e.g. 'for slide in slides:').
        """
        for sldId in self._sldIdLst:
            yield self.part.related_slide(sldId.rId)

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slides) == 4').
        """
        return len(self._sldIdLst)

    def add_slide(self, slide_layout):
        """
        Return a newly added slide that inherits layout from *slide_layout*.
        """
        rId, slide = self.part.add_slide(slide_layout)
        slide.shapes.clone_layout_placeholders(slide_layout)
        self._sldIdLst.add_sldId(rId)
        return slide

    def get(self, slide_id, default=None):
        """
        Return the slide identified by integer *slide_id* in this
        presentation, or *default* if not found.
        """
        slide = self.part.get_slide(slide_id)
        if slide is None:
            return default
        return slide

    def index(self, slide):
        """
        Map *slide* to an integer representing its zero-based position in
        this slide collection. Raises |ValueError| on *slide* not present.
        """
        for idx, this_slide in enumerate(self):
            if this_slide == slide:
                return idx
        raise ValueError('%s is not in slide collection' % slide)


class SlideLayout(_BaseSlide):
    """
    Slide layout object. Provides access to placeholders, regular shapes, and
    slide layout-level properties.
    """

    __slots__ = ('_placeholders', '_shapes')

    def iter_cloneable_placeholders(self):
        """
        Generate a reference to each layout placeholder on this slide layout
        that should be cloned to a slide when the layout is applied to that
        slide.
        """
        latent_ph_types = (
            PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.SLIDE_NUMBER
        )
        for ph in self.placeholders:
            if ph.ph_type not in latent_ph_types:
                yield ph

    @lazyproperty
    def placeholders(self):
        """
        Instance of |LayoutPlaceholders| containing sequence of placeholder
        shapes in this slide layout, sorted in *idx* order.
        """
        return LayoutPlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self):
        """
        Instance of |LayoutShapes| containing the sequence of shapes
        appearing on this slide layout.
        """
        return LayoutShapes(self._element.spTree, self)

    @property
    def slide_master(self):
        """
        Slide master from which this slide layout inherits properties.
        """
        return self.part.slide_master


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
        return self.part.related_slide_layout(sldLayoutId.rId)

    def __iter__(self):
        """
        Generate a reference to each of the |SlideLayout| instances in the
        collection, in sequence.
        """
        for sldLayoutId in self._sldLayoutIdLst:
            yield self.part.related_slide_layout(sldLayoutId.rId)

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slides) == 4').
        """
        return len(self._sldLayoutIdLst)


class SlideMaster(_BaseSlide):
    """
    Slide master object. Provides access to placeholders, regular shapes,
    slide layouts, and slide master-level properties.
    """

    __slots__ = ('_placeholders', '_shapes', '_slide_layouts')

    @lazyproperty
    def placeholders(self):
        """
        Instance of |MasterPlaceholders| containing sequence of placeholder
        shapes in this slide master, sorted in *idx* order.
        """
        return MasterPlaceholders(self._element.spTree, self)

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
        return self.part.related_slide_master(sldMasterId.rId)

    def __iter__(self):
        """
        Generate a reference to each of the |SlideMaster| instances in the
        collection, in sequence.
        """
        for smi in self._sldMasterIdLst:
            yield self.part.related_slide_master(smi.rId)

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slide_masters) == 4').
        """
        return len(self._sldMasterIdLst)
