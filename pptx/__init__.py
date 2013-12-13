# encoding: utf-8

"""
Initialization module for python-pptx
"""

__version__ = '0.3.0'


import pptx.exc as exceptions
import sys
sys.modules['pptx.exceptions'] = exceptions
del sys

from pptx.api import Presentation  # noqa

from pptx.parts.coreprops import CoreProperties
from pptx.parts.image import Image
from pptx.parts.slides import Slide, SlideLayout, SlideMaster
from pptx.presentation import Presentation as _Presentation

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import PartFactory

content_type_to_part_class_map = {
    CT.PML_PRESENTATION_MAIN: _Presentation,
    CT.PML_TEMPLATE_MAIN:     _Presentation,
    CT.PML_SLIDESHOW_MAIN:    _Presentation,
    CT.OPC_CORE_PROPERTIES:   CoreProperties,
    CT.PML_SLIDE:             Slide,
    CT.PML_SLIDE_LAYOUT:      SlideLayout,
    CT.PML_SLIDE_MASTER:      SlideMaster,
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
    CoreProperties, Image, Slide, SlideLayout, SlideMaster, _Presentation,
    CT, PartFactory
)
