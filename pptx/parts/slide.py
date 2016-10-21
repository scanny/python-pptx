# encoding: utf-8

"""
Slide and related objects.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .chart import ChartPart
from ..opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from ..opc.package import XmlPart
from ..opc.packuri import PackURI
from ..oxml.slide import CT_NotesMaster, CT_Slide
from ..slide import NotesMaster, Slide, SlideLayout, SlideMaster
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


class NotesMasterPart(BaseSlidePart):
    """
    Notes master part. Corresponds to package file
    `ppt/notesMasters/notesMaster1.xml`.
    """
    @classmethod
    def create_default(cls, package):
        """
        Create and return a default notes master part, including creating the
        new theme it requires.
        """
        notes_master_part = cls._new(package)
        theme_part = cls._new_theme_part(package)
        notes_master_part.relate_to(theme_part, RT.THEME)
        return notes_master_part

    @lazyproperty
    def notes_master(self):
        """
        Return the |NotesMaster| object that proxies this notes master part.
        """
        return NotesMaster(self._element, self)

    @classmethod
    def _new(cls, package):
        """
        Create and return a standalone, default notes master part based on
        the built-in template (without any related parts, such as theme).
        """
        partname = PackURI('/ppt/notesMasters/notesMaster1.xml')
        content_type = CT.PML_NOTES_MASTER
        notesMaster = CT_NotesMaster.new_default()
        return NotesMasterPart(partname, content_type, notesMaster, package)

    @classmethod
    def _new_theme_part(cls, package):
        """
        Create and return a default theme part suitable for use with a notes
        master.
        """
        raise NotImplementedError


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
        contained in this part.
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
    def slide_id(self):
        """
        Return the slide identifier stored in the presentation part for this
        slide part.
        """
        presentation_part = self.package.presentation_part
        return presentation_part.slide_id(self)

    @property
    def slide_layout(self):
        """
        |SlideLayout| object the slide in this part inherits from.
        """
        slide_layout_part = self.part_related_by(RT.SLIDE_LAYOUT)
        return slide_layout_part.slide_layout


class SlideLayoutPart(BaseSlidePart):
    """
    Slide layout part. Corresponds to package files
    ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """
    @lazyproperty
    def slide_layout(self):
        """
        The |SlideLayout| object representing this part.
        """
        return SlideLayout(self._element, self)

    @property
    def slide_master(self):
        """
        Slide master from which this slide layout inherits properties.
        """
        return self.part_related_by(RT.SLIDE_MASTER).slide_master


class SlideMasterPart(BaseSlidePart):
    """
    Slide master part. Corresponds to package files
    ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    """
    def related_slide_layout(self, rId):
        """
        Return the |SlideLayout| object of the related |SlideLayoutPart|
        corresponding to relationship key *rId*.
        """
        return self.related_parts[rId].slide_layout

    @lazyproperty
    def slide_master(self):
        """
        The |SlideMaster| object representing this part.
        """
        return SlideMaster(self._element, self)
