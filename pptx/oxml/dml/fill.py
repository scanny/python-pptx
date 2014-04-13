# encoding: utf-8

"""
lxml custom element classes for DrawingML-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from ..ns import qn
from ..shared import SubElement


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


class CT_BlipFillProperties(OxmlElement):
    """
    Custom element class for <a:blipFill> element.
    """


class CT_GradientFillProperties(OxmlElement):
    """
    Custom element class for <a:gradFill> element.
    """


class CT_GroupFillProperties(OxmlElement):
    """
    Custom element class for <a:grpFill> element.
    """


class CT_NoFillProperties(OxmlElement):
    """
    Custom element class for <a:NoFill> element.
    """


class CT_PatternFillProperties(OxmlElement):
    """
    Custom element class for <a:pattFill> element.
    """


class CT_SolidColorFillProperties(OxmlElement):
    """
    Custom element class for <a:solidFill> element.
    """
    @property
    def eg_colorchoice(self):
        """
        Return the child representing the EG_ColorChoice element group in
        this element, or |None| if no such child is present.
        """
        return self._first_child_found_in(
            'a:scrgbClr', 'a:srgbClr', 'a:hslClr', 'a:sysClr', 'a:schemeClr',
            'a:prstClr'
        )

    def get_or_change_to_schemeClr(self):
        """
        Return the <a:schemeClr> child of this <a:solidFill>, replacing any
        other EG_ColorChoice element if found, perhaps most commonly a
        <a:srgbClr> element.
        """
        if self.schemeClr is not None:
            return self.schemeClr
        self._clear_color_choice()
        return SubElement(self, 'a:schemeClr')

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


class EG_FillProperties(object):

    __member_names__ = (
        'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill',
        'a:grpFill'
    )
