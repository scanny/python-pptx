# encoding: utf-8

"""
Slide and related objects.
"""

from __future__ import absolute_import

from warnings import warn

from .chart import ChartPart
from ..opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from ..opc.package import XmlPart
from ..oxml.parts.slide import CT_Slide
from ..shapes.shapetree import SlideShapeFactory, SlideShapeTree
from ..shared import ParentedElementProxy
from ..util import lazyproperty


class BaseSlide(XmlPart):
    """
    Base class for slide parts, e.g. slide, slideLayout, slideMaster,
    notesSlide, notesMaster, and handoutMaster.
    """
    def get_image(self, rId):
        """
        Return an |Image| object containing the image related to this slide
        by *rId*. Raises |KeyError| if no image is related by that id, which
        would generally indicate a corrupted .pptx file.
        """
        return self.related_parts[rId].image

    def get_or_add_image_part(self, image_file):
        """
        Return an ``(image_part, rId)`` 2-tuple corresponding to an
        |ImagePart| object containing the image in *image_file*, and related
        to this slide with the key *rId*. If either the image part or
        relationship already exists, they are reused, otherwise they are
        newly created.
        """
        image_part = self._package.get_or_add_image_part(image_file)
        rId = self.relate_to(image_part, RT.IMAGE)
        return image_part, rId

    @property
    def name(self):
        """
        Internal name of this slide.
        """
        return self._element.cSld.name

    @property
    def spTree(self):
        """
        Reference to ``<p:spTree>`` element for this slide
        """
        return self._element.cSld.spTree


class Slide(BaseSlide):
    """
    Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml.
    """
    @classmethod
    def new(cls, slide_layout, partname, package):
        """
        Return a new slide based on *slide_layout* and having *partname*,
        created from scratch.
        """
        slide_elm = CT_Slide.new()
        slide = cls(partname, CT.PML_SLIDE, slide_elm, package)
        slide.shapes.clone_layout_placeholders(slide_layout)
        slide.relate_to(slide_layout, RT.SLIDE_LAYOUT)
        return slide

    def add_chart_part(self, chart_type, chart_data):
        """
        Return the rId of a new |ChartPart| object containing a chart of
        *chart_type*, displaying *chart_data*, and related to the slide
        containing this shape tree.
        """
        chart_part = ChartPart.new(chart_type, chart_data, self.package)
        rId = self.relate_to(chart_part, RT.CHART)
        return rId

    @lazyproperty
    def placeholders(self):
        """
        Instance of |_SlidePlaceholders| containing sequence of placeholder
        shapes in this slide.
        """
        return _SlidePlaceholders(self._element.spTree, self)

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
        return self.part_related_by(RT.SLIDE_LAYOUT)

    @property
    def slidelayout(self):
        """
        Deprecated. Use ``.slide_layout`` property instead.
        """
        msg = (
            'Slide.slidelayout property is deprecated. Use .slide_layout '
            'instead.'
        )
        warn(msg, UserWarning, stacklevel=2)
        return self.slide_layout


class _SlidePlaceholders(ParentedElementProxy):
    """
    Collection of placeholder shapes on a slide. Supports iteration,
    :func:`len`, and dictionary-style lookup on the `idx` value of the
    placeholders it contains.
    """

    __slots__ = ()

    def __getitem__(self, idx):
        """
        Access placeholder shape having *idx*. Note that while this looks
        like list access, idx is actually a dictionary key and will raise
        |KeyError| if no placeholder with that idx value is in the
        collection.
        """
        for e in self._element.iter_ph_elms():
            if e.ph_idx == idx:
                return SlideShapeFactory(e, self)
        raise KeyError('no placeholder on this slide with idx == %d' % idx)

    def __iter__(self):
        """
        Generate placeholder shapes in `idx` order.
        """
        ph_elms = sorted(
            [e for e in self._element.iter_ph_elms()], key=lambda e: e.ph_idx
        )
        return (SlideShapeFactory(e, self) for e in ph_elms)

    def __len__(self):
        """
        Return count of placeholder shapes.
        """
        return len(list(self._element.iter_ph_elms()))
