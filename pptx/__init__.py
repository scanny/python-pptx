# encoding: utf-8

"""
Initialization module for python-pptx
"""

__version__ = '0.5.8'


import pptx.exc as exceptions
import sys
sys.modules['pptx.exceptions'] = exceptions
del sys

from pptx.api import Presentation  # noqa

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import PartFactory
from pptx.parts.chart import ChartPart
from pptx.parts.coreprops import CoreProperties
from pptx.parts.image import ImagePart
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import Slide
from pptx.parts.slidelayout import SlideLayout
from pptx.parts.slidemaster import SlideMaster

content_type_to_part_class_map = {
    CT.PML_PRESENTATION_MAIN: PresentationPart,
    CT.PML_PRES_MACRO_MAIN:   PresentationPart,
    CT.PML_TEMPLATE_MAIN:     PresentationPart,
    CT.PML_SLIDESHOW_MAIN:    PresentationPart,
    CT.OPC_CORE_PROPERTIES:   CoreProperties,
    CT.PML_SLIDE:             Slide,
    CT.PML_SLIDE_LAYOUT:      SlideLayout,
    CT.PML_SLIDE_MASTER:      SlideMaster,
    CT.DML_CHART:             ChartPart,
    CT.BMP:                   ImagePart,
    CT.GIF:                   ImagePart,
    CT.JPEG:                  ImagePart,
    CT.MS_PHOTO:              ImagePart,
    CT.PNG:                   ImagePart,
    CT.TIFF:                  ImagePart,
    CT.X_EMF:                 ImagePart,
    CT.X_WMF:                 ImagePart,
}

PartFactory.part_type_for.update(content_type_to_part_class_map)

del (
    ChartPart, CoreProperties, ImagePart, Slide, SlideLayout, SlideMaster,
    PresentationPart, CT, PartFactory
)
