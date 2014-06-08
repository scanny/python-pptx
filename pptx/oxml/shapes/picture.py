# encoding: utf-8

"""
lxml custom element classes for picture-related XML elements.
"""

from __future__ import absolute_import

from .. import parse_xml
from ..ns import nsdecls
from .shared import BaseShapeElement
from ..xmlchemy import BaseOxmlElement, OneAndOnlyOne


class CT_Picture(BaseShapeElement):
    """
    ``<p:pic>`` element, which represents a picture shape (an image placement
    on a slide).
    """
    nvPicPr = OneAndOnlyOne('p:nvPicPr')
    spPr = OneAndOnlyOne('p:spPr')

    def get_or_add_ln(self):
        """
        Return the <a:ln> grandchild element, newly added if not present.
        """
        return self.spPr.get_or_add_ln()

    @property
    def ln(self):
        """
        ``<a:ln>`` grand-child element or |None| if not present
        """
        return self.spPr.ln

    @classmethod
    def new_pic(cls, id_, name, desc, rId, left, top, width, height):
        """
        Return a new ``<p:pic>`` element tree configured with the supplied
        parameters.
        """
        xml = cls._pic_tmpl() % (
            id_, name, desc, rId, left, top, width, height
        )
        pic = parse_xml(xml)
        return pic

    @classmethod
    def _pic_tmpl(cls):
        return (
            '<p:pic %s>\n'
            '  <p:nvPicPr>\n'
            '    <p:cNvPr id="%s" name="%s" descr="%s"/>\n'
            '    <p:cNvPicPr>\n'
            '      <a:picLocks noChangeAspect="1"/>\n'
            '    </p:cNvPicPr>\n'
            '    <p:nvPr/>\n'
            '  </p:nvPicPr>\n'
            '  <p:blipFill>\n'
            '    <a:blip r:embed="%s"/>\n'
            '    <a:stretch>\n'
            '      <a:fillRect/>\n'
            '    </a:stretch>\n'
            '  </p:blipFill>\n'
            '  <p:spPr>\n'
            '    <a:xfrm>\n'
            '      <a:off x="%s" y="%s"/>\n'
            '      <a:ext cx="%s" cy="%s"/>\n'
            '    </a:xfrm>\n'
            '    <a:prstGeom prst="rect">\n'
            '      <a:avLst/>\n'
            '    </a:prstGeom>\n'
            '  </p:spPr>\n'
            '</p:pic>' % (
                nsdecls('a', 'p', 'r'), '%d', '%s', '%s', '%s', '%d', '%d',
                '%d', '%d'
            )
        )


class CT_PictureNonVisual(BaseOxmlElement):
    """
    ``<p:nvPicPr>`` element, containing non-visual properties for a picture
    shape.
    """
    cNvPr = OneAndOnlyOne('p:cNvPr')
