# encoding: utf-8

"""
lxml custom element classes for text-related XML elements.
"""

from __future__ import absolute_import

from . import parse_xml
from ..enum.text import MSO_AUTO_SIZE
from .ns import nsdecls, qn
from .shared import Element, SubElement
from .simpletypes import (
    ST_Coordinate32, ST_TextFontSize, ST_TextTypeface, XsdBoolean, XsdString
)
from .xmlchemy import (
    BaseOxmlElement, Choice, OneAndOnlyOne, OneOrMore, OptionalAttribute,
    RequiredAttribute, ZeroOrOne, ZeroOrOneChoice
)


class CT_Hyperlink(BaseOxmlElement):
    """
    Custom element class for <a:hlinkClick> elements.
    """
    rId = OptionalAttribute('r:id', XsdString)


class CT_RegularTextRun(BaseOxmlElement):
    """
    Custom element class for <a:r> elements.
    """
    rPr = ZeroOrOne('a:rPr', successors=('a:t',))
    t = OneAndOnlyOne('a:t')


class CT_TextBody(BaseOxmlElement):
    """
    <p:txBody> custom element class
    """
    bodyPr = OneAndOnlyOne('a:bodyPr')
    p = OneOrMore('a:p')

    @classmethod
    def new(cls):
        """
        Return a new ``<p:txBody>`` element tree
        """
        xml = cls._txBody_tmpl()
        txBody = parse_xml(xml)
        return txBody

    @classmethod
    def new_a_txBody(cls):
        """
        Return a new ``<a:txBody>`` element tree, suitable for use in a table
        cell and possibly other situations.
        """
        xml = cls._a_txBody_tmpl()
        txBody = parse_xml(xml)
        return txBody

    @classmethod
    def _a_txBody_tmpl(cls):
        return (
            '<a:txBody %s>\n'
            '  <a:bodyPr/>\n'
            '  <a:p/>\n'
            '</a:txBody>\n' % (nsdecls('a'))
        )

    @classmethod
    def _txBody_tmpl(cls):
        return (
            '<p:txBody %s>\n'
            '  <a:bodyPr/>\n'
            '  <a:lstStyle/>\n'
            '  <a:p/>\n'
            '</p:txBody>\n' % (nsdecls('a', 'p'))
        )


class CT_TextBodyProperties(BaseOxmlElement):
    """
    <a:bodyPr> custom element class
    """
    eg_textAutoFit = ZeroOrOneChoice((
        Choice('a:noAutofit'), Choice('a:normAutofit'),
        Choice('a:spAutoFit')),
        successors=('a:scene3d', 'a:sp3d', 'a:flatTx', 'a:extLst')
    )
    lIns = OptionalAttribute('lIns', ST_Coordinate32)
    tIns = OptionalAttribute('tIns', ST_Coordinate32)
    rIns = OptionalAttribute('rIns', ST_Coordinate32)
    bIns = OptionalAttribute('bIns', ST_Coordinate32)

    @property
    def autofit(self):
        """
        The autofit setting for the textframe, a member of the
        ``MSO_AUTO_SIZE`` enumeration.
        """
        if self.noAutofit is not None:
            return MSO_AUTO_SIZE.NONE
        if self.normAutofit is not None:
            return MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        if self.spAutoFit is not None:
            return MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        return None

    @autofit.setter
    def autofit(self, value):
        if value is not None and not MSO_AUTO_SIZE.validate(value):
            raise ValueError(
                'only None or a member of the MSO_AUTO_SIZE enumeration can '
                'be assigned to CT_TextBodyProperties.autofit, got %s'
                % value
            )
        self._remove_eg_textAutoFit()
        if value == MSO_AUTO_SIZE.NONE:
            self._add_noAutofit()
        elif value == MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE:
            self._add_normAutofit()
        elif value == MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT:
            self._add_spAutoFit()


class CT_TextCharacterProperties(BaseOxmlElement):
    """
    Custom element class for all of <a:rPr>, <a:defRPr>, and <a:endParaRPr>
    elements. 'rPr' is short for 'run properties', and it corresponds to the
    _Font proxy class.
    """
    eg_fillProperties = ZeroOrOneChoice((
        Choice('a:noFill'), Choice('a:solidFill'), Choice('a:gradFill'),
        Choice('a:blipFill'), Choice('a:pattFill'), Choice('a:grpFill')),
        successors=(
            'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn',
            'a:uFillTx', 'a:uFill', 'a:latin', 'a:ea', 'a:cs', 'a:sym',
            'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'
        )
    )
    latin = ZeroOrOne('a:latin', successors=(
        'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl',
        'a:extLst'
    ))
    hlinkClick = ZeroOrOne('a:hlinkClick', successors=(
        'a:hlinkMouseOver', 'a:rtl', 'a:extLst'
    ))

    sz = OptionalAttribute('sz', ST_TextFontSize)
    b = OptionalAttribute('b', XsdBoolean)
    i = OptionalAttribute('i', XsdBoolean)

    def add_hlinkClick(self, rId):
        """
        Add an <a:hlinkClick> child element with r:id attribute set to *rId*.
        """
        hlinkClick = self.get_or_add_hlinkClick()
        hlinkClick.set(qn('r:id'), rId)
        return hlinkClick


class CT_TextFont(BaseOxmlElement):
    """
    Custom element class for <a:latin>, <a:ea>, <a:cs>, and <a:sym> child
    elements of CT_TextCharacterProperties, e.g. <a:rPr>.
    """
    typeface = RequiredAttribute('typeface', ST_TextTypeface)


class CT_TextParagraph(BaseOxmlElement):
    """
    <a:p> custom element class
    """
    pPr = ZeroOrOne('a:pPr', successors=(
        'a:r', 'a:br', 'a:fld', 'a:endParaRPr'
    ))
    endParaRPr = ZeroOrOne('a:endParaRPr', successors=())

    def add_r(self):
        """
        Return a newly appended <a:r> element.
        """
        r = Element('a:r')
        SubElement(r, 'a:t')
        # work out where to insert it, ahead of a:endParaRPr if there is one
        try:
            self.endParaRPr.addprevious(r)
        except AttributeError:
            self.append(r)
        return r

    def remove_child_r_elms(self):
        """
        Return self after removing all <a:r> child elements.
        """
        children = self.getchildren()
        for child in children:
            if child.tag == qn('a:r'):
                self.remove(child)
        return self


class CT_TextParagraphProperties(BaseOxmlElement):
    """
    <a:pPr> custom element class
    """
    defRPr = ZeroOrOne('a:defRPr', successors=('a:extLst',))

    @property
    def algn(self):
        """
        Paragraph horizontal alignment value, like ``TAT.CENTER``. Value of
        'algn' attribute on <a:pPr> child element. None if no 'algn'
        attribute is present.
        """
        return self.get('algn')

    @algn.setter
    def algn(self, value):
        if value is None:
            if 'algn' in self.attrib:
                del self.attrib['algn']
            return
        self.set('algn', value)
