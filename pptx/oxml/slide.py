# encoding: utf-8

"""
lxml custom element classes for slide master-related XML elements.
"""

from __future__ import absolute_import

from .xmlchemy import BaseOxmlElement, OneAndOnlyOne, ZeroOrOne


class CT_CommonSlideData(BaseOxmlElement):
    """
    ``<p:cSld>`` element.
    """
    spTree = OneAndOnlyOne('p:spTree')


class CT_Slide(BaseOxmlElement):
    """
    ``<p:sld>`` element, root of a slide part
    """
    cSld = OneAndOnlyOne('p:cSld')
    clrMapOvr = ZeroOrOne('p:clrMapOvr', successors=(
        'p:transition', 'p:timing', 'p:extLst'
    ))
