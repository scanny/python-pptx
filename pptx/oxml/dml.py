# encoding: utf-8

"""
lxml custom element classes for DrawingML-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml.ns import qn


class OxmlElement(objectify.ObjectifiedElement):
    pass


class CT_Percentage(OxmlElement):
    """
    Custom element class for <a:lumMod> and <a:lumOff> elements.
    """
    @property
    def val(self):
        return self.get('val')


class CT_SchemeColor(OxmlElement):
    """
    Custom element class for <a:schemeClr> element.
    """
    @property
    def lumMod(self):
        """
        The <a:lumMod> child element, or None if not present.
        """
        return self.find(qn('a:lumMod'))

    @property
    def lumOff(self):
        """
        The <a:lumOff> child element, or None if not present.
        """
        return self.find(qn('a:lumOff'))

    @property
    def val(self):
        return self.get('val')


class CT_SRgbColor(OxmlElement):
    """
    Custom element class for <a:srgbClr> element.
    """
    @property
    def lumMod(self):
        """
        The <a:lumMod> child element, or None if not present.
        """
        return self.find(qn('a:lumMod'))

    @property
    def lumOff(self):
        """
        The <a:lumOff> child element, or None if not present.
        """
        return self.find(qn('a:lumOff'))

    @property
    def val(self):
        return self.get('val')


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
