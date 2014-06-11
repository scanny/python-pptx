# encoding: utf-8

"""
lxml custom element classes for slide master-related XML elements.
"""

from __future__ import absolute_import

from . import parse_xml
from .ns import nsdecls
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

    @classmethod
    def new(cls):
        """
        Return a new ``<p:sld>`` element configured as a base slide shape.
        """
        return parse_xml(cls._sld_xml())

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
