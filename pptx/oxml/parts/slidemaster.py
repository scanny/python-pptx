# encoding: utf-8

"""
lxml custom element classes for slide master-related XML elements.
"""

from __future__ import absolute_import

from ..simpletypes import XsdString
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrOne, ZeroOrMore
)


class CT_SlideLayoutIdList(BaseOxmlElement):
    """
    ``<p:sldLayoutIdLst>`` element, child of ``<p:sldMaster>`` containing
    references to the slide layouts that inherit from the slide master.
    """
    sldLayoutId = ZeroOrMore('p:sldLayoutId')


class CT_SlideLayoutIdListEntry(BaseOxmlElement):
    """
    ``<p:sldLayoutId>`` element, child of ``<p:sldLayoutIdLst>`` containing
    a reference to a slide layout.
    """
    rId = RequiredAttribute('r:id', XsdString)


class CT_SlideMaster(BaseOxmlElement):
    """
    ``<p:sldMaster>`` element, root of a slide master part
    """
    cSld = OneAndOnlyOne('p:cSld')
    sldLayoutIdLst = ZeroOrOne('p:sldLayoutIdLst', successors=(
        'p:transition', 'p:timing', 'p:hf', 'p:txStyles', 'p:extLst'
    ))
