# encoding: utf-8

"""
lxml custom element classes for DrawingML-related XML elements.
"""

from __future__ import absolute_import

from ...enum.dml import MSO_THEME_COLOR
from ..ns import qn
from ..shared import SubElement
from ..xmlchemy import BaseOxmlElement


class _BaseColorElement(BaseOxmlElement):
    """
    Base class for <a:srgbClr> and <a:schemeClr> elements.
    """
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

    @val.setter
    def val(self, value):
        self.set('val', value)


class CT_HslColor(_BaseColorElement):
    """
    Custom element class for <a:hslClr> element.
    """


class CT_Percentage(BaseOxmlElement):
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
    @property
    def val(self):
        val = self.get('val')
        mso_theme_color_idx = MSO_THEME_COLOR.from_xml(val)
        return mso_theme_color_idx

    @val.setter
    def val(self, mso_theme_color_idx):
        val = MSO_THEME_COLOR.to_xml(mso_theme_color_idx)
        self.set('val', val)


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
