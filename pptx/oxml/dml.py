# encoding: utf-8

"""
lxml custom element classes for DrawingML-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml.core import SubElement
from pptx.oxml.ns import qn


class OxmlElement(objectify.ObjectifiedElement):
    pass


class _BaseColorElement(OxmlElement):
    """
    Base class for <a:srgbClr> and <a:schemeClr> elements.
    """
    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in ('val',):
            self.set(name, value)
        else:
            super(CT_SchemeColor, self).__setattr__(name, value)

    def add_lumMod(self, value):
        """
        Return a newly added <a:lumMod> child element.
        """
        return SubElement(self, 'a:lumMod', val=str(value))

    def add_lumOff(self, value):
        """
        Return a newly added <a:lumOff> child element.
        """
        return SubElement(self, 'a:lumOff', val=str(value))

    def clear_lum(self):
        """
        Return self after removing any <a:lumMod> and <a:lumOff> child
        elements.
        """
        lum_tagnames = (qn('a:lumMod'), qn('a:lumOff'))
        for child in self.getchildren():
            if child.tag in lum_tagnames:
                self.remove(child)
        return self

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


class CT_Percentage(OxmlElement):
    """
    Custom element class for <a:lumMod> and <a:lumOff> elements.
    """
    @property
    def val(self):
        return self.get('val')


class CT_SchemeColor(_BaseColorElement):
    """
    Custom element class for <a:schemeClr> element.
    """


class CT_SRgbColor(_BaseColorElement):
    """
    Custom element class for <a:srgbClr> element.
    """


class CT_SolidColorFillProperties(OxmlElement):
    """
    Custom element class for <a:solidFill> element.
    """
    def get_or_change_to_schemeClr(self):
        """
        Return the <a:schemeClr> child of this <a:solidFill>, replacing any
        other EG_ColorChoice element if found, perhaps most commonly a
        <a:srgbClr> element.
        """
        if self.schemeClr is not None:
            return self.schemeClr
        self._clear_color_choice()
        return self._add_schemeClr()

    def get_or_change_to_srgbClr(self):
        """
        Return the <a:srgbClr> child of this <a:solidFill>, replacing any
        other EG_ColorChoice element if found, perhaps most commonly a
        <a:schemeClr> element.
        """
        if self.srgbClr is not None:
            return self.srgbClr
        self._clear_color_choice()
        return self._add_srgbClr()

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

    def _add_schemeClr(self):
        """
        Return a newly added <a:schemeClr> child element.
        """
        return SubElement(self, 'a:schemeClr')

    def _add_srgbClr(self):
        """
        Return a newly added <a:srgbClr> child element.
        """
        return SubElement(self, 'a:srgbClr')

    def _clear_color_choice(self):
        """
        Remove the EG_ColorChoice child element, e.g. <a:schemeClr>.
        """
        eg_colorchoice_tagnames = (
            'a:scrgbClr', 'a:srgbClr', 'a:hslClr', 'a:sysClr', 'a:schemeClr',
            'a:prstClr'
        )
        for tagname in eg_colorchoice_tagnames:
            element = self.find(qn(tagname))
            if element is not None:
                self.remove(element)
