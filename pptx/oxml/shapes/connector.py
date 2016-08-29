# encoding: utf-8

"""
lxml custom element classes for shape-related XML elements.
"""

from __future__ import absolute_import

from .. import parse_xml
from ..ns import nsdecls
from .shared import BaseShapeElement
from ..xmlchemy import BaseOxmlElement, OneAndOnlyOne


class CT_Connector(BaseShapeElement):
    """
    A line/connector shape ``<p:cxnSp>`` element
    """
    _tag_seq = (
        'p:nvCxnSpPr', 'p:spPr', 'p:style', 'p:extLst'
    )
    spPr = OneAndOnlyOne('p:spPr')
    del _tag_seq

    @classmethod
    def new_cxnSp(cls, id_, name, prst, x, y, cx, cy, flipH, flipV):
        """
        Return a new ``<p:cxnSp>`` element tree configured as a base
        connector.
        """
        tmpl = cls._cxnSp_tmpl()
        flip = (
            (' flipH="1"' if flipH else '') + (' flipV="1"' if flipV else '')
        )
        xml = tmpl.format(**{
            'nsdecls': nsdecls('a', 'p'),
            'id':      id_,
            'name':    name,
            'x':       x,
            'y':       y,
            'cx':      cx,
            'cy':      cy,
            'prst':    prst,
            'flip':    flip,
        })
        return parse_xml(xml)

    @staticmethod
    def _cxnSp_tmpl():
        return (
            '<p:cxnSp {nsdecls}>\n'
            '  <p:nvCxnSpPr>\n'
            '    <p:cNvPr id="{id}" name="{name}"/>\n'
            '    <p:cNvCxnSpPr/>\n'
            '    <p:nvPr/>\n'
            '  </p:nvCxnSpPr>\n'
            '  <p:spPr>\n'
            '    <a:xfrm{flip}>\n'
            '      <a:off x="{x}" y="{y}"/>\n'
            '      <a:ext cx="{cx}" cy="{cy}"/>\n'
            '    </a:xfrm>\n'
            '    <a:prstGeom prst="{prst}">\n'
            '      <a:avLst/>\n'
            '    </a:prstGeom>\n'
            '  </p:spPr>\n'
            '  <p:style>\n'
            '    <a:lnRef idx="2">\n'
            '      <a:schemeClr val="accent1"/>\n'
            '    </a:lnRef>\n'
            '    <a:fillRef idx="0">\n'
            '      <a:schemeClr val="accent1"/>\n'
            '    </a:fillRef>\n'
            '    <a:effectRef idx="1">\n'
            '      <a:schemeClr val="accent1"/>\n'
            '    </a:effectRef>\n'
            '    <a:fontRef idx="minor">\n'
            '      <a:schemeClr val="tx1"/>\n'
            '    </a:fontRef>\n'
            '  </p:style>\n'
            '</p:cxnSp>'
        )


class CT_ConnectorNonVisual(BaseOxmlElement):
    """
    ``<p:nvCxnSpPr>`` element, container for the non-visual properties of
    a connector, such as name, id, etc.
    """
    cNvPr = OneAndOnlyOne('p:cNvPr')
