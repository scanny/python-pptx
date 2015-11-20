# encoding: utf-8

"""
lxml custom element classes for DrawingML-related XML elements.
"""

from __future__ import absolute_import

from ..simpletypes import ST_Percentage, ST_RelationshipId
from ..xmlchemy import (
    BaseOxmlElement, Choice, OptionalAttribute, ZeroOrOne, ZeroOrOneChoice
)


class CT_Blip(BaseOxmlElement):
    """
    <a:blip> element
    """
    rEmbed = OptionalAttribute('r:embed', ST_RelationshipId)


class CT_BlipFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:blipFill> element.
    """
    _tag_seq = ('a:blip', 'a:srcRect', 'a:tile', 'a:stretch')
    blip = ZeroOrOne('a:blip', successors=_tag_seq[1:])
    srcRect = ZeroOrOne('a:srcRect', successors=_tag_seq[2:])
    del _tag_seq

    def crop(self, cropping):
        """
        Set `a:srcRect` child to crop according to *cropping* values.
        """
        srcRect = self._add_srcRect()
        srcRect.l, srcRect.t, srcRect.r, srcRect.b = cropping


class CT_GradientFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:gradFill> element.
    """


class CT_GroupFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:grpFill> element.
    """


class CT_NoFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:NoFill> element.
    """


class CT_PatternFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:pattFill> element.
    """


class CT_RelativeRect(BaseOxmlElement):
    """
    `a:srcRect` element and perhaps others.
    """
    l = OptionalAttribute('l', ST_Percentage, default=0.0)
    t = OptionalAttribute('t', ST_Percentage, default=0.0)
    r = OptionalAttribute('r', ST_Percentage, default=0.0)
    b = OptionalAttribute('b', ST_Percentage, default=0.0)


class CT_SolidColorFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:solidFill> element.
    """
    eg_colorChoice = ZeroOrOneChoice((
        Choice('a:scrgbClr'), Choice('a:srgbClr'), Choice('a:hslClr'),
        Choice('a:sysClr'), Choice('a:schemeClr'), Choice('a:prstClr')),
        successors=()
    )
