# encoding: utf-8

"""
DrawingML objects related to fill, FillFormat being the top-most.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..enum.dml import MSO_FILL
from ..oxml.dml.fill import (
    CT_BlipFillProperties, CT_GradientFillProperties, CT_GroupFillProperties,
    CT_NoFillProperties, CT_PatternFillProperties,
    CT_SolidColorFillProperties,
)
from pptx.util import lazyproperty

from .color import ColorFormat


class FillFormat(object):
    """
    Provides access to the current fill properties object and provides
    methods to change the fill type.
    """
    def __init__(self, eg_fill_properties_parent, fill_obj):
        super(FillFormat, self).__init__()
        self._xPr = eg_fill_properties_parent
        self._fill = fill_obj

    def background(self):
        """
        Sets the fill type to noFill, i.e. transparent.
        """
        noFill = self._xPr.get_or_change_to_noFill()
        self._fill = _NoFill(noFill)

    @property
    def fore_color(self):
        """
        Return a |ColorFormat| instance representing the foreground color of
        this fill.
        """
        return self._fill.fore_color

    @classmethod
    def from_fill_parent(cls, eg_fillProperties_parent):
        """
        Return a |FillFormat| instance initialized to the settings contained
        in *eg_fillProperties_parent*, which must be an element having
        EG_FillProperties in its child element sequence in the XML schema.
        """
        fill_elm = eg_fillProperties_parent.eg_fillProperties
        fill = _Fill(fill_elm)
        fill_format = cls(eg_fillProperties_parent, fill)
        return fill_format

    def solid(self):
        """
        Sets the fill type to solid, i.e. a solid color. Note that calling
        this method does not set a color or by itself cause the shape to
        appear with a solid color fill; rather it enables subsequent
        assignments to properties like fore_color to set the color.
        """
        solidFill = self._xPr.get_or_change_to_solidFill()
        self._fill = _SolidFill(solidFill)

    @property
    def type(self):
        """
        Return a value from the :ref:`MsoFillType` enumeration corresponding
        to the type of this fill.
        """
        return self._fill.type


class _Fill(object):
    """
    Object factory for fill object of class matching fill element, such as
    _SolidFill for ``<a:solidFill>``; also serves as the base class for all
    fill classes
    """
    def __new__(cls, xFill):
        if xFill is None:
            fill_cls = _NoneFill
        elif isinstance(xFill, CT_BlipFillProperties):
            fill_cls = _BlipFill
        elif isinstance(xFill, CT_GradientFillProperties):
            fill_cls = _GradFill
        elif isinstance(xFill, CT_GroupFillProperties):
            fill_cls = _GrpFill
        elif isinstance(xFill, CT_NoFillProperties):
            fill_cls = _NoFill
        elif isinstance(xFill, CT_PatternFillProperties):
            fill_cls = _PattFill
        elif isinstance(xFill, CT_SolidColorFillProperties):
            fill_cls = _SolidFill
        else:
            fill_cls = _Fill
        return super(_Fill, cls).__new__(fill_cls)

    @property
    def fore_color(self):
        """
        Raise NotImplementedError for all fill types that are still skeleton
        subclasses.
        """
        tmpl = ".fore_color property not implemented yet for %s"
        raise NotImplementedError(tmpl % self.__class__.__name__)

    @property
    def type(self):  # pragma: no cover
        tmpl = ".type property must be implemented on %s"
        raise NotImplementedError(tmpl % self.__class__.__name__)


class _BlipFill(_Fill):

    @property
    def fore_color(self):
        """
        Raise TypeError with message explaining why this doesn't make sense.
        """
        tmpl = "a picture fill has no foreground color"
        raise TypeError(tmpl)

    @property
    def type(self):
        return MSO_FILL.PICTURE


class _GradFill(_Fill):

    @property
    def type(self):
        return MSO_FILL.GRADIENT


class _GrpFill(_Fill):

    @property
    def fore_color(self):
        """
        Raise TypeError with message explaining why this doesn't make sense.
        """
        tmpl = "a group fill has no foreground color"
        raise TypeError(tmpl)

    @property
    def type(self):
        return MSO_FILL.GROUP


class _NoFill(_Fill):

    @property
    def fore_color(self):
        """
        Raise TypeError with message explaining why this doesn't make sense.
        """
        tmpl = "a transparent (background) fill has no foreground color"
        raise TypeError(tmpl)

    @property
    def type(self):
        return MSO_FILL.BACKGROUND


class _NoneFill(_Fill):

    @property
    def fore_color(self):
        """
        Raise TypeError with message explaining why this doesn't make sense.
        """
        tmpl = "can't set .fore_color on no fill, call .solid() first"
        raise TypeError(tmpl)

    @property
    def type(self):
        return None


class _PattFill(_Fill):

    @property
    def type(self):
        return MSO_FILL.PATTERNED


class _SolidFill(_Fill):
    """
    Provides access to fill properties such as color for solid fills.
    """
    def __init__(self, solidFill):
        super(_SolidFill, self).__init__()
        self._solidFill = solidFill

    @lazyproperty
    def fore_color(self):
        return ColorFormat.from_colorchoice_parent(self._solidFill)

    @property
    def type(self):
        return MSO_FILL.SOLID
