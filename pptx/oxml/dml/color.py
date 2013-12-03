# encoding: utf-8

"""
lxml custom element classes for DrawingML-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml.core import SubElement
from pptx.oxml.ns import qn


class OxmlElement(objectify.ObjectifiedElement):

    def _first_child_found_in(self, *tagnames):
        """
        Return the first child found with tag in *tagnames*, or None if
        not found.
        """
        for tagname in tagnames:
            child = self.find(qn(tagname))
            if child is not None:
                return child
        return None


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


class CT_HslColor(_BaseColorElement):
    """
    Custom element class for <a:hslClr> element.
    """


class CT_Percentage(OxmlElement):
    """
    Custom element class for <a:lumMod> and <a:lumOff> elements.
    """
    @property
    def val(self):
        return self.get('val')


class CT_PresetColor(_BaseColorElement):
    """
    Custom element class for <a:prstClr> element.
    """


class CT_SchemeColor(_BaseColorElement):
    """
    Custom element class for <a:schemeClr> element.
    """


class CT_ScRgbColor(_BaseColorElement):
    """
    Custom element class for <a:scrgbClr> element.
    """


class CT_SRgbColor(_BaseColorElement):
    """
    Custom element class for <a:srgbClr> element.
    """


class CT_SystemColor(_BaseColorElement):
    """
    Custom element class for <a:sysClr> element.
    """
