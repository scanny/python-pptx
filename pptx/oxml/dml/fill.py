# encoding: utf-8

"""
lxml custom element classes for DrawingML-related XML elements.
"""

from __future__ import absolute_import

from pptx.enum.dml import MSO_PATTERN_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import ST_Percentage, ST_RelationshipId
from pptx.oxml.xmlchemy import (
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
    """`a:pattFill` custom element class"""
    _tag_seq = ('a:fgClr', 'a:bgClr')
    fgClr = ZeroOrOne('a:fgClr', successors=_tag_seq[1:])
    bgClr = ZeroOrOne('a:bgClr', successors=_tag_seq[2:])
    del _tag_seq
    prst = OptionalAttribute('prst', MSO_PATTERN_TYPE)

    def _new_bgClr(self):
        """Override default to add minimum subtree."""
        xml = (
            '<a:bgClr %s>\n'
            ' <a:srgbClr val="FFFFFF"/>\n'
            '</a:bgClr>\n'
        ) % nsdecls('a')
        bgClr = parse_xml(xml)
        return bgClr

    def _new_fgClr(self):
        """Override default to add minimum subtree."""
        xml = (
            '<a:fgClr %s>\n'
            ' <a:srgbClr val="000000"/>\n'
            '</a:fgClr>\n'
        ) % nsdecls('a')
        fgClr = parse_xml(xml)
        return fgClr


class CT_RelativeRect(BaseOxmlElement):
    """`a:srcRect` element and perhaps others."""
    l = OptionalAttribute('l', ST_Percentage, default=0.0)
    t = OptionalAttribute('t', ST_Percentage, default=0.0)
    r = OptionalAttribute('r', ST_Percentage, default=0.0)
    b = OptionalAttribute('b', ST_Percentage, default=0.0)


class CT_SolidColorFillProperties(BaseOxmlElement):
    """`a:solidFill` custom element class."""
    eg_colorChoice = ZeroOrOneChoice((
        Choice('a:scrgbClr'), Choice('a:srgbClr'), Choice('a:hslClr'),
        Choice('a:sysClr'), Choice('a:schemeClr'), Choice('a:prstClr')),
        successors=()
    )
