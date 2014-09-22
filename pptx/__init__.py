# encoding: utf-8

"""
Initialization module for python-pptx
"""

__version__ = '0.5.1'


import pptx.exc as exceptions
import sys
sys.modules['pptx.exceptions'] = exceptions
del sys

from pptx.api import Presentation  # noqa

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import PartFactory
from pptx.parts.chart import ChartPart
from pptx.parts.coreprops import CoreProperties
from pptx.parts.image import Image
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import Slide
from pptx.parts.slidelayout import SlideLayout
from pptx.parts.slidemaster import SlideMaster

content_type_to_part_class_map = {
    CT.PML_PRESENTATION_MAIN: PresentationPart,
    CT.PML_TEMPLATE_MAIN:     PresentationPart,
    CT.PML_SLIDESHOW_MAIN:    PresentationPart,
    CT.OPC_CORE_PROPERTIES:   CoreProperties,
    CT.PML_SLIDE:             Slide,
    CT.PML_SLIDE_LAYOUT:      SlideLayout,
    CT.PML_SLIDE_MASTER:      SlideMaster,
    CT.DML_CHART:             ChartPart,
    CT.BMP:                   Image,
    CT.GIF:                   Image,
    CT.JPEG:                  Image,
    CT.MS_PHOTO:              Image,
    CT.PNG:                   Image,
    CT.TIFF:                  Image,
    CT.X_EMF:                 Image,
    CT.X_WMF:                 Image,
}

PartFactory.part_type_for.update(content_type_to_part_class_map)

del (
    ChartPart, CoreProperties, Image, Slide, SlideLayout, SlideMaster,
    PresentationPart, CT, PartFactory
)
