# encoding: utf-8

"""
lxml custom element classes for slide master-related XML elements.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from . import parse_xml
from .ns import nsdecls
from .simpletypes import XsdString
from .xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, RequiredAttribute,
    ZeroOrMore, ZeroOrOne
)


class CT_CommonSlideData(BaseOxmlElement):
    """
    ``<p:cSld>`` element.
    """
    _tag_seq = (
        'p:bg', 'p:spTree', 'p:custDataLst', 'p:controls', 'p:extLst'
    )
    spTree = OneAndOnlyOne('p:spTree')
    del _tag_seq

    name = OptionalAttribute('name', XsdString, default='')


class CT_Slide(BaseOxmlElement):
    """
    ``<p:sld>`` element, root of a slide part
    """
    _tag_seq = (
        'p:cSld', 'p:clrMapOvr', 'p:transition', 'p:timing', 'p:extLst'
    )
    cSld = OneAndOnlyOne('p:cSld')
    clrMapOvr = ZeroOrOne('p:clrMapOvr', successors=_tag_seq[2:])
    del _tag_seq

    @classmethod
    def new(cls):
        """
        Return a new ``<p:sld>`` element configured as a base slide shape.
        """
        return parse_xml(cls._sld_xml())

    @property
    def spTree(self):
        """
        Return required `p:cSld/p:spTree` grandchild.
        """
        return self.cSld.spTree

    @staticmethod
    def _sld_xml():
        return (
            '<p:sld %s>\n'
            '  <p:cSld>\n'
            '    <p:spTree>\n'
            '      <p:nvGrpSpPr>\n'
            '        <p:cNvPr id="1" name=""/>\n'
            '        <p:cNvGrpSpPr/>\n'
            '        <p:nvPr/>\n'
            '      </p:nvGrpSpPr>\n'
            '      <p:grpSpPr/>\n'
            '    </p:spTree>\n'
            '  </p:cSld>\n'
            '  <p:clrMapOvr>\n'
            '    <a:masterClrMapping/>\n'
            '  </p:clrMapOvr>\n'
            '</p:sld>' % nsdecls('a', 'p', 'r')
        )


class CT_SlideLayout(BaseOxmlElement):
    """
    ``<p:sldLayout>`` element, root of a slide layout part
    """
    _tag_seq = (
        'p:cSld', 'p:clrMapOvr', 'p:transition', 'p:timing', 'p:hf',
        'p:extLst'
    )
    cSld = OneAndOnlyOne('p:cSld')
    del _tag_seq

    @property
    def spTree(self):
        """
        Return required `p:cSld/p:spTree` grandchild.
        """
        return self.cSld.spTree


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
    _tag_seq = (
        'p:cSld', 'p:clrMap', 'p:sldLayoutIdLst', 'p:transition', 'p:timing',
        'p:hf', 'p:txStyles', 'p:extLst'
    )
    cSld = OneAndOnlyOne('p:cSld')
    sldLayoutIdLst = ZeroOrOne('p:sldLayoutIdLst', successors=_tag_seq[3:])
    del _tag_seq

    @property
    def spTree(self):
        """
        Return required `p:cSld/p:spTree` grandchild.
        """
        return self.cSld.spTree
