# encoding: utf-8

"""
lxml custom element classes for text-related XML elements.
"""

from __future__ import absolute_import

from . import parse_xml
from ..enum.text import (
    MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, PP_PARAGRAPH_ALIGNMENT
)
from .ns import nsdecls
from .simpletypes import (
    ST_Coordinate32, ST_TextFontSize, ST_TextIndentLevelType,
    ST_TextTypeface, ST_TextWrappingType, XsdBoolean, XsdString
)
from .xmlchemy import (
    BaseOxmlElement, Choice, OneAndOnlyOne, OneOrMore, OptionalAttribute,
    RequiredAttribute, ZeroOrMore, ZeroOrOne, ZeroOrOneChoice
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
    ``<p:txBody>`` custom element class, also used for ``<c:txPr>`` in
    charts and perhaps other elements.
    """
    bodyPr = OneAndOnlyOne('a:bodyPr')
    p = OneOrMore('a:p')

    @property
    def defRPr(self):
        """
        ``<a:defRPr>`` element of required first ``p`` child, added with its
        ancestors if not present. Used when element is a ``<c:txPr>`` in
        a chart and the ``p`` element is used only to specify formatting, not
        content.
        """
        p = self.p_lst[0]
        pPr = p.get_or_add_pPr()
        defRPr = pPr.get_or_add_defRPr()
        return defRPr

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
    def new_txPr(cls):
        """
        Return a ``<c:txPr>`` element tree suitable for use in a chart object
        like data labels or tick labels.
        """
        xml = (
            '<c:txPr %s>\n'
            '  <a:bodyPr/>\n'
            '  <a:lstStyle/>\n'
            '  <a:p>\n'
            '    <a:pPr>\n'
            '      <a:defRPr/>\n'
            '    </a:pPr>\n'
            '  </a:p>\n'
            '</c:txPr>\n'
        ) % nsdecls('c', 'a')
        txPr = parse_xml(xml)
        return txPr

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
    anchor = OptionalAttribute('anchor', MSO_VERTICAL_ANCHOR)
    wrap = OptionalAttribute('wrap', ST_TextWrappingType)

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
        if value is not None and value not in MSO_AUTO_SIZE._valid_settings:
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
    |Font| proxy class.
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
        hlinkClick.rId = rId
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
    r = ZeroOrMore('a:r', successors=('a:endParaRPr',))
    endParaRPr = ZeroOrOne('a:endParaRPr', successors=())

    def add_r(self):
        """
        Return a newly appended <a:r> element.
        """
        return self._add_r()

    def remove_child_r_elms(self):
        """
        Return self after removing all <a:r> child elements.
        """
        for r in self.r_lst:
            self.remove(r)
        return self

    def _new_r(self):
        r_xml = '<a:r %s><a:t/></a:r>' % nsdecls('a')
        return parse_xml(r_xml)


class CT_TextParagraphProperties(BaseOxmlElement):
    """
    <a:pPr> custom element class
    """
    defRPr = ZeroOrOne('a:defRPr', successors=('a:extLst',))
    lvl = OptionalAttribute('lvl', ST_TextIndentLevelType, default=0)
    algn = OptionalAttribute('algn', PP_PARAGRAPH_ALIGNMENT)
