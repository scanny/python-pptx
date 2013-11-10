# encoding: utf-8

"""
lxml custom element classes for DrawingML-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml.ns import qn


class OxmlElement(objectify.ObjectifiedElement):
    pass


class CT_SolidColorFillProperties(OxmlElement):
    """
    Custom element class for <a:solidFill> element.
    """
    @property
    def schemeClr(self):
        """
        The <a:schemeClr> child element, or None if not present.
        """
        return self.find(qn('a:schemeClr'))

    @property
    def srgbClr(self):
        """
        The <a:srgbClr> child element, or None if not present.
        """
        return self.find(qn('a:srgbClr'))
