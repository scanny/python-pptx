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
from ..shared import ParentedElementProxy


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


class SlidePlaceholders(ParentedElementProxy):
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
