# encoding: utf-8

"""
Slide and related objects.
"""

from __future__ import absolute_import

from .chart import ChartPart
from ..opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from ..opc.package import XmlPart
from ..oxml.parts.slide import CT_Slide
from ..slide import Slide
from ..util import lazyproperty


class BaseSlidePart(XmlPart):
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


class SlidePart(BaseSlidePart):
    """
    Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml.
    """
    @classmethod
    def new(cls, partname, package, slide_layout_part):
        """
        Return a newly-created blank slide part having *partname* and related
        to *slide_layout_part*.
        """
        sld = CT_Slide.new()
        slide_part = cls(partname, CT.PML_SLIDE, sld, package)
        slide_part.relate_to(slide_layout_part, RT.SLIDE_LAYOUT)
        return slide_part

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
    def slide(self):
        """
        The |Slide| object representing this slide part.
        """
        return Slide(self._element, self)

    @property
    def slide_layout(self):
        """
        |SlideLayout| object the slide in this part inherits from.
        """
        slide_layout_part = self.part_related_by(RT.SLIDE_LAYOUT)
        return slide_layout_part.slide_layout
