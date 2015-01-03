# encoding: utf-8

"""
Factory functions for making the right shape types from shape elemens.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .autoshape import Shape
from .base import BaseShape
from .graphfrm import GraphicFrame
from ..oxml.ns import qn
from .picture import Picture
from .placeholder import SlidePlaceholder


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
    tag_name = shape_elm.tag
    if tag_name == qn('p:sp') and shape_elm.has_ph_elm:
        return SlidePlaceholder(shape_elm, parent)
    return BaseShapeFactory(shape_elm, parent)
