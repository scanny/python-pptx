# encoding: utf-8

"""
lxml custom element classes for shape-related XML elements.
"""

from __future__ import absolute_import

from .. import parse_xml
from ...enum.shapes import MSO_AUTO_SHAPE_TYPE, PP_PLACEHOLDER
from ..ns import nsdecls
from .shared import BaseShapeElement
from ..simpletypes import XsdBoolean, XsdString
from ..text import CT_TextBody
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, RequiredAttribute,
    ZeroOrOne, ZeroOrMore
)


class CT_GeomGuide(BaseOxmlElement):
    """
    ``<a:gd>`` custom element class, defining a "guide", corresponding to
    a yellow diamond-shaped handle on an autoshape.
    """
    name = RequiredAttribute('name', XsdString)
    fmla = RequiredAttribute('fmla', XsdString)


class CT_GeomGuideList(BaseOxmlElement):
    """
    ``<a:avLst>`` custom element class
    """
    gd = ZeroOrMore('a:gd')


class CT_NonVisualDrawingShapeProps(BaseShapeElement):
    """
    ``<p:cNvSpPr>`` custom element class
    """
    spLocks = ZeroOrOne('a:spLocks')
    txBox = OptionalAttribute('txBox', XsdBoolean)


class CT_PresetGeometry2D(BaseOxmlElement):
    """
    <a:prstGeom> custom element class
    """
    avLst = ZeroOrOne('a:avLst')
    prst = RequiredAttribute('prst', MSO_AUTO_SHAPE_TYPE)

    @property
    def gd_lst(self):
        """
        Sequence containing the ``gd`` element children of ``<a:avLst>``
        child element, empty if none are present.
        """
        avLst = self.avLst
        if avLst is None:
            return []
        return avLst.gd_lst

    def rewrite_guides(self, guides):
        """
        Remove any ``<a:gd>`` element children of ``<a:avLst>`` and replace
        them with ones having (name, val) in *guides*.
        """
        self._remove_avLst()
        avLst = self._add_avLst()
        for name, val in guides:
            gd = avLst._add_gd()
            gd.name = name
            gd.fmla = 'val %d' % val


class CT_Shape(BaseShapeElement):
    """
    ``<p:sp>`` custom element class
    """
    nvSpPr = OneAndOnlyOne('p:nvSpPr')
    spPr = OneAndOnlyOne('p:spPr')
    txBody = ZeroOrOne('p:txBody', successors=('p:extLst',))

    def get_or_add_ln(self):
        """
        Return the <a:ln> grandchild element, newly added if not present.
        """
        return self.spPr.get_or_add_ln()

    @property
    def is_autoshape(self):
        """
        True if this shape is an auto shape. A shape is an auto shape if it
        has a ``<a:prstGeom>`` element and does not have a txBox="1" attribute
        on cNvSpPr.
        """
        prstGeom = self.prstGeom
        if prstGeom is None:
            return False
        if self.nvSpPr.cNvSpPr.txBox is True:
            return False
        return True

    @property
    def is_textbox(self):
        """
        True if this shape is a text box. A shape is a text box if it has a
        ``txBox`` attribute on cNvSpPr that resolves to |True|. The default
        when the txBox attribute is missing is |False|.
        """
        if self.nvSpPr.cNvSpPr.txBox is True:
            return True
        return False

    @property
    def ln(self):
        """
        ``<a:ln>`` grand-child element or |None| if not present
        """
        return self.spPr.ln

    @staticmethod
    def new_autoshape_sp(id_, name, prst, left, top, width, height):
        """
        Return a new ``<p:sp>`` element tree configured as a base auto shape.
        """
        tmpl = CT_Shape._autoshape_sp_tmpl()
        xml = tmpl % (id_, name, left, top, width, height, prst)
        sp = parse_xml(xml)
        return sp

    @staticmethod
    def new_placeholder_sp(id_, name, ph_type, orient, sz, idx):
        """
        Return a new ``<p:sp>`` element tree configured as a placeholder
        shape.
        """
        tmpl = CT_Shape._ph_sp_tmpl()
        xml = tmpl % (id_, name)
        sp = parse_xml(xml)

        ph = sp.nvSpPr.nvPr.get_or_add_ph()
        ph.type = ph_type
        ph.idx = idx
        ph.orient = orient
        ph.sz = sz

        placeholder_types_that_have_a_text_frame = (
            PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE,
            PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.OBJECT
        )

        if ph_type in placeholder_types_that_have_a_text_frame:
            sp.append(CT_TextBody.new())

        return sp

    @staticmethod
    def new_textbox_sp(id_, name, left, top, width, height):
        """
        Return a new ``<p:sp>`` element tree configured as a base textbox
        shape.
        """
        tmpl = CT_Shape._textbox_sp_tmpl()
        xml = tmpl % (id_, name, left, top, width, height)
        sp = parse_xml(xml)
        return sp

    @property
    def prst(self):
        """
        Value of ``prst`` attribute of ``<a:prstGeom>`` element or |None| if
        not present.
        """
        prstGeom = self.prstGeom
        if prstGeom is None:
            return None
        return prstGeom.prst

    @property
    def prstGeom(self):
        """
        Reference to ``<a:prstGeom>`` child element or |None| if this shape
        doesn't have one, for example, if it's a placeholder shape.
        """
        return self.spPr.prstGeom

    @staticmethod
    def _autoshape_sp_tmpl():
        return (
            '<p:sp %s>\n'
            '  <p:nvSpPr>\n'
            '    <p:cNvPr id="%s" name="%s"/>\n'
            '    <p:cNvSpPr/>\n'
            '    <p:nvPr/>\n'
            '  </p:nvSpPr>\n'
            '  <p:spPr>\n'
            '    <a:xfrm>\n'
            '      <a:off x="%s" y="%s"/>\n'
            '      <a:ext cx="%s" cy="%s"/>\n'
            '    </a:xfrm>\n'
            '    <a:prstGeom prst="%s">\n'
            '      <a:avLst/>\n'
            '    </a:prstGeom>\n'
            '  </p:spPr>\n'
            '  <p:style>\n'
            '    <a:lnRef idx="1">\n'
            '      <a:schemeClr val="accent1"/>\n'
            '    </a:lnRef>\n'
            '    <a:fillRef idx="3">\n'
            '      <a:schemeClr val="accent1"/>\n'
            '    </a:fillRef>\n'
            '    <a:effectRef idx="2">\n'
            '      <a:schemeClr val="accent1"/>\n'
            '    </a:effectRef>\n'
            '    <a:fontRef idx="minor">\n'
            '      <a:schemeClr val="lt1"/>\n'
            '    </a:fontRef>\n'
            '  </p:style>\n'
            '  <p:txBody>\n'
            '    <a:bodyPr rtlCol="0" anchor="ctr"/>\n'
            '    <a:lstStyle/>\n'
            '    <a:p>\n'
            '      <a:pPr algn="ctr"/>\n'
            '    </a:p>\n'
            '  </p:txBody>\n'
            '</p:sp>' %
            (nsdecls('a', 'p'), '%d', '%s', '%d', '%d', '%d', '%d', '%s')
        )

    def _new_txBody(self):
        return CT_TextBody.new_p_txBody()

    @staticmethod
    def _ph_sp_tmpl():
        return (
            '<p:sp %s>\n'
            '  <p:nvSpPr>\n'
            '    <p:cNvPr id="%s" name="%s"/>\n'
            '    <p:cNvSpPr>\n'
            '      <a:spLocks noGrp="1"/>\n'
            '    </p:cNvSpPr>\n'
            '    <p:nvPr/>\n'
            '  </p:nvSpPr>\n'
            '  <p:spPr/>\n'
            '</p:sp>' % (nsdecls('a', 'p'), '%d', '%s')
        )

    @staticmethod
    def _textbox_sp_tmpl():
        return (
            '<p:sp %s>\n'
            '  <p:nvSpPr>\n'
            '    <p:cNvPr id="%s" name="%s"/>\n'
            '    <p:cNvSpPr txBox="1"/>\n'
            '    <p:nvPr/>\n'
            '  </p:nvSpPr>\n'
            '  <p:spPr>\n'
            '    <a:xfrm>\n'
            '      <a:off x="%s" y="%s"/>\n'
            '      <a:ext cx="%s" cy="%s"/>\n'
            '    </a:xfrm>\n'
            '    <a:prstGeom prst="rect">\n'
            '      <a:avLst/>\n'
            '    </a:prstGeom>\n'
            '    <a:noFill/>\n'
            '  </p:spPr>\n'
            '  <p:txBody>\n'
            '    <a:bodyPr wrap="none">\n'
            '      <a:spAutoFit/>\n'
            '    </a:bodyPr>\n'
            '    <a:lstStyle/>\n'
            '    <a:p/>\n'
            '  </p:txBody>\n'
            '</p:sp>' %
            (nsdecls('a', 'p'), '%d', '%s', '%d', '%d', '%d', '%d')
        )


class CT_ShapeNonVisual(BaseShapeElement):
    """
    ``<p:nvSpPr>`` custom element class
    """
    cNvPr = OneAndOnlyOne('p:cNvPr')
    cNvSpPr = OneAndOnlyOne('p:cNvSpPr')
    nvPr = OneAndOnlyOne('p:nvPr')
