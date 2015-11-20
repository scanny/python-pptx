# encoding: utf-8

"""
Factory functions for making the right shape types from shape elemens.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .autoshape import Shape
from .base import BaseShape
from ..enum.shapes import PP_PLACEHOLDER
from .graphfrm import GraphicFrame
from ..oxml.ns import qn
from .picture import Picture
from .placeholder import (
    ChartPlaceholder, PicturePlaceholder, PlaceholderGraphicFrame,
    PlaceholderPicture, SlidePlaceholder, TablePlaceholder
)


def BaseShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*.
    """
    tag_name = shape_elm.tag
    if tag_name == qn('p:sp'):
        return Shape(shape_elm, parent)
    if tag_name == qn('p:pic'):
        return Picture(shape_elm, parent)
    if tag_name == qn('p:graphicFrame'):
        return GraphicFrame(shape_elm, parent)
    return BaseShape(shape_elm, parent)


def SlideShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*
    on a slide.
    """
    if shape_elm.has_ph_elm:
        return _SlidePlaceholderFactory(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)


def _SlidePlaceholderFactory(shape_elm, parent):
    """
    Return a placeholder shape of the appropriate type for *shape_elm*.
    """
    tag = shape_elm.tag
    if tag == qn('p:sp'):
        Constructor = {
            PP_PLACEHOLDER.BITMAP:  PicturePlaceholder,
            PP_PLACEHOLDER.CHART:   ChartPlaceholder,
            PP_PLACEHOLDER.PICTURE: PicturePlaceholder,
            PP_PLACEHOLDER.TABLE:   TablePlaceholder,
        }.get(shape_elm.ph_type, SlidePlaceholder)
    elif tag == qn('p:graphicFrame'):
        Constructor = PlaceholderGraphicFrame
    elif tag == qn('p:pic'):
        Constructor = PlaceholderPicture
    else:
        Constructor = BaseShapeFactory
    return Constructor(shape_elm, parent)
