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
from ..oxml.slide import CT_NotesMaster, CT_NotesSlide, CT_Slide
from ..oxml.theme import CT_OfficeStyleSheet
from ..slide import NotesMaster, NotesSlide, Slide, SlideLayout, SlideMaster
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
        partname = package.next_partname('/ppt/theme/theme%d.xml')
        content_type = CT.OFC_THEME
        theme = CT_OfficeStyleSheet.new_default()
        return XmlPart(partname, content_type, theme, package)


class NotesSlidePart(BaseSlidePart):
    """
    Notes slide part. Contains the slide notes content and the layout for the
    slide handout page. Corresponds to package file
    `ppt/notesSlides/notesSlide[1-9][0-9]*.xml`.
    """
    @classmethod
    def new(cls, package, slide_part):
        """
        Create and return a new notes slide part based on the notes master
        and related to both the notes master part and *slide_part*. If no
        notes master is present, create one based on the default template.
        """
        notes_master_part = package.presentation_part.notes_master_part
        notes_slide_part = cls._add_notes_slide_part(
            package, slide_part, notes_master_part
        )
        notes_slide = notes_slide_part.notes_slide
        notes_slide.clone_master_placeholders(notes_master_part.notes_master)
        return notes_slide_part

    @lazyproperty
    def notes_master(self):
        """
        Return the |NotesMaster| object this notes slide inherits from.
        """
        notes_master_part = self.part_related_by(RT.NOTES_MASTER)
        return notes_master_part.notes_master

    @lazyproperty
    def notes_slide(self):
        """
        Return the |NotesSlide| object that proxies this notes slide part.
        """
        return NotesSlide(self._element, self)

    @classmethod
    def _add_notes_slide_part(cls, package, slide_part, notes_master_part):
        """
        Create and return a new notes slide part that is fully related, but
        has no shape content (i.e. placeholders not cloned).
        """
        partname = package.next_partname('/ppt/notesSlides/notesSlide%d.xml')
        content_type = CT.PML_NOTES_SLIDE
        notes = CT_NotesSlide.new()
        notes_slide_part = NotesSlidePart(
            partname, content_type, notes, package
        )
        notes_slide_part.relate_to(notes_master_part, RT.NOTES_MASTER)
        notes_slide_part.relate_to(slide_part, RT.SLIDE)
        return notes_slide_part


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

    def get_or_add_video_media_part(self, video):
        """Return rIds for media and video relationships to media part.

        A new |MediaPart| object is created if it does not already exist
        (such as would occur if the same video appeared more than once in
         a presentation). Two relationships to the media part are created,
        one each with MEDIA and VIDEO relationship types. The need for two
        appears to be for legacy support for an earlier (pre-Office 2010)
        PowerPoint media embedding strategy.
        """
        media_part = self._package.get_or_add_media_part(video)
        media_rId = self.relate_to(media_part, RT.MEDIA)
        video_rId = self.relate_to(media_part, RT.VIDEO)
        return media_rId, video_rId

    @property
    def has_notes_slide(self):
        """
        Return True if this slide has a notes slide, False otherwise. A notes
        slide is created by the :attr:`notes_slide` property when one doesn't
        exist; use this property to test for a notes slide without the
        possible side-effect of creating one.
        """
        try:
            self.part_related_by(RT.NOTES_SLIDE)
        except KeyError:
            return False
        return True

    @lazyproperty
    def notes_slide(self):
        """
        The |NotesSlide| instance associated with this slide. If the slide
        does not have a notes slide, a new one is created. The same single
        instance is returned on each call.
        """
        try:
            notes_slide_part = self.part_related_by(RT.NOTES_SLIDE)
        except KeyError:
            notes_slide_part = self._add_notes_slide_part()
        return notes_slide_part.notes_slide

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

    def _add_notes_slide_part(self):
        """
        Return a newly created |NotesSlidePart| object related to this slide
        part. Caller is responsible for ensuring this slide doesn't already
        have a notes slide part.
        """
        notes_slide_part = NotesSlidePart.new(self.package, self)
        self.relate_to(notes_slide_part, RT.NOTES_SLIDE)
        return notes_slide_part


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
