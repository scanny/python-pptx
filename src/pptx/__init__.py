"""Initialization module for python-pptx package."""

from __future__ import annotations

import sys
from typing import TYPE_CHECKING

import pptx.exc as exceptions
from pptx.api import Presentation
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import PartFactory
from pptx.parts.chart import ChartPart
from pptx.parts.chartdrawing import ChartDrawingPart
from pptx.parts.comments import CommentAuthorsPart, CommentsPart
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.customprops import CustomPropertiesPart
from pptx.parts.extprops import ExtendedPropertiesPart
from pptx.parts.font import FontPart
from pptx.parts.image import ImagePart
from pptx.parts.media import MediaPart
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import (
    NotesMasterPart,
    NotesSlidePart,
    SlideLayoutPart,
    SlideMasterPart,
    SlidePart,
    ThemePart,
)
from pptx.parts.tags import TagsPart
from pptx.parts.viewprops import ViewPropsPart

if TYPE_CHECKING:
    from pptx.opc.package import Part

__version__ = "2026.05.0"

sys.modules["pptx.exceptions"] = exceptions
del sys

__all__ = ["Presentation"]

content_type_to_part_class_map: dict[str, type[Part]] = {
    CT.PML_PRESENTATION_MAIN: PresentationPart,
    CT.PML_PRES_MACRO_MAIN: PresentationPart,
    CT.PML_TEMPLATE_MAIN: PresentationPart,
    CT.PML_SLIDESHOW_MAIN: PresentationPart,
    CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
    CT.OFC_CUSTOM_PROPERTIES: CustomPropertiesPart,
    CT.OFC_EXTENDED_PROPERTIES: ExtendedPropertiesPart,
    CT.PML_COMMENTS: CommentsPart,
    CT.PML_COMMENT_AUTHORS: CommentAuthorsPart,
    CT.PML_NOTES_MASTER: NotesMasterPart,
    CT.PML_NOTES_SLIDE: NotesSlidePart,
    CT.PML_SLIDE: SlidePart,
    CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
    CT.PML_SLIDE_MASTER: SlideMasterPart,
    CT.PML_TAGS: TagsPart,
    CT.PML_VIEW_PROPS: ViewPropsPart,
    CT.OFC_THEME: ThemePart,
    CT.DML_CHART: ChartPart,
    CT.DML_CHARTSHAPES: ChartDrawingPart,
    CT.X_FONTDATA: FontPart,
    CT.X_FONT_TTF: FontPart,
    CT.BMP: ImagePart,
    CT.GIF: ImagePart,
    CT.JPEG: ImagePart,
    CT.MS_PHOTO: ImagePart,
    CT.PNG: ImagePart,
    CT.TIFF: ImagePart,
    CT.X_EMF: ImagePart,
    CT.X_WMF: ImagePart,
    CT.ASF: MediaPart,
    CT.AUDIO: MediaPart,
    CT.AUDIO_AIFF: MediaPart,
    CT.AUDIO_MIDI: MediaPart,
    CT.AUDIO_MP3: MediaPart,
    CT.AUDIO_MP4: MediaPart,
    CT.AUDIO_MPEG: MediaPart,
    CT.AUDIO_OGG: MediaPart,
    CT.AUDIO_WAV: MediaPart,
    CT.AUDIO_X_MS_WMA: MediaPart,
    CT.AUDIO_X_WAV: MediaPart,
    CT.AVI: MediaPart,
    CT.MOV: MediaPart,
    CT.MP4: MediaPart,
    CT.MPG: MediaPart,
    CT.MS_VIDEO: MediaPart,
    CT.SWF: MediaPart,
    CT.VIDEO: MediaPart,
    CT.WAV: MediaPart,
    CT.WMV: MediaPart,
    CT.X_MS_VIDEO: MediaPart,
    # -- accommodate "image/jpg" as an alias for "image/jpeg" --
    "image/jpg": ImagePart,
    # -- accommodate "image/tif" as an alias for "image/tiff" (issue #1084) --
    "image/tif": ImagePart,
}

PartFactory.part_type_for.update(content_type_to_part_class_map)

del (
    ChartPart,
    ChartDrawingPart,
    CommentAuthorsPart,
    CommentsPart,
    CorePropertiesPart,
    CustomPropertiesPart,
    ExtendedPropertiesPart,
    FontPart,
    ImagePart,
    MediaPart,
    SlidePart,
    SlideLayoutPart,
    SlideMasterPart,
    TagsPart,
    ThemePart,
    ViewPropsPart,
    PresentationPart,
    CT,
    PartFactory,
)
