# encoding: utf-8

"""
lxml custom element classes for picture-related XML elements.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

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
    blipFill = OneAndOnlyOne('p:blipFill')
    spPr = OneAndOnlyOne('p:spPr')

    def crop_to_fit(self, image_size, view_size):
        """
        Set cropping values in `p:blipFill/a:srcRect` such that an image of
        *image_size* will stretch to exactly fit *view_size* when its aspect
        ratio is preserved.
        """
        self.blipFill.crop(self._fill_cropping(image_size, view_size))

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
    def new_ph_pic(cls, id_, name, desc, rId):
        """
        Return a new `p:pic` placeholder element populated with the supplied
        parameters.
        """
        return parse_xml(
            cls._pic_ph_tmpl() % (id_, name, desc, rId)
        )

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

    def _fill_cropping(self, image_size, view_size):
        """
        Return a (left, top, right, bottom) 4-tuple containing the cropping
        values required to display an image of *image_size* in *view_size*
        when stretched proportionately. Each value is a percentage expressed
        as a fraction of 1.0, e.g. 0.425 represents 42.5%. *image_size* and
        *view_size* are each (width, height) pairs.
        """
        def aspect_ratio(width, height):
            return width / height

        ar_view = aspect_ratio(*view_size)
        ar_image = aspect_ratio(*image_size)

        if ar_view < ar_image:  # image too wide
            crop = (1.0 - (ar_view/ar_image)) / 2.0
            return (crop, 0.0, crop, 0.0)
        if ar_view > ar_image:  # image too tall
            crop = (1.0 - (ar_image/ar_view)) / 2.0
            return (0.0, crop, 0.0, crop)
        return (0.0, 0.0, 0.0, 0.0)

    @classmethod
    def _pic_ph_tmpl(cls):
        return (
            '<p:pic %s>\n'
            '  <p:nvPicPr>\n'
            '    <p:cNvPr id="%%d" name="%%s" descr="%%s"/>\n'
            '    <p:cNvPicPr>\n'
            '      <a:picLocks noGrp="1" noChangeAspect="1"/>\n'
            '    </p:cNvPicPr>\n'
            '    <p:nvPr/>\n'
            '  </p:nvPicPr>\n'
            '  <p:blipFill>\n'
            '    <a:blip r:embed="%%s"/>\n'
            '    <a:stretch>\n'
            '      <a:fillRect/>\n'
            '    </a:stretch>\n'
            '  </p:blipFill>\n'
            '  <p:spPr/>\n'
            '</p:pic>' % nsdecls('p', 'a', 'r')
        )

    @classmethod
    def _pic_tmpl(cls):
        return (
            '<p:pic %s>\n'
            '  <p:nvPicPr>\n'
            '    <p:cNvPr id="%%d" name="%%s" descr="%%s"/>\n'
            '    <p:cNvPicPr>\n'
            '      <a:picLocks noChangeAspect="1"/>\n'
            '    </p:cNvPicPr>\n'
            '    <p:nvPr/>\n'
            '  </p:nvPicPr>\n'
            '  <p:blipFill>\n'
            '    <a:blip r:embed="%%s"/>\n'
            '    <a:stretch>\n'
            '      <a:fillRect/>\n'
            '    </a:stretch>\n'
            '  </p:blipFill>\n'
            '  <p:spPr>\n'
            '    <a:xfrm>\n'
            '      <a:off x="%%d" y="%%d"/>\n'
            '      <a:ext cx="%%d" cy="%%d"/>\n'
            '    </a:xfrm>\n'
            '    <a:prstGeom prst="rect">\n'
            '      <a:avLst/>\n'
            '    </a:prstGeom>\n'
            '  </p:spPr>\n'
            '</p:pic>' % nsdecls('a', 'p', 'r')
        )


class CT_PictureNonVisual(BaseOxmlElement):
    """
    ``<p:nvPicPr>`` element, containing non-visual properties for a picture
    shape.
    """
    cNvPr = OneAndOnlyOne('p:cNvPr')
