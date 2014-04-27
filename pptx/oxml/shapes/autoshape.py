# encoding: utf-8

"""
lxml custom element classes for shape-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from .. import parse_xml_bytes
from ..ns import nsdecls, qn
from .shared import (
    BaseShapeElement, ST_Direction, ST_PlaceholderSize, ST_PlaceholderType
)
from ..shared import BaseOxmlElement, child, SubElement
from ..text import CT_TextBody


class CT_PresetGeometry2D(BaseOxmlElement):
    """<a:prstGeom> custom element class"""
    @property
    def gd(self):
        """
        Sequence containing the ``gd`` element children of ``<a:avLst>``
        child element, empty if none are present.
        """
        try:
            gd_elms = tuple([gd for gd in self.avLst.gd])
        except AttributeError:
            gd_elms = ()
        return gd_elms

    @property
    def prst(self):
        """Value of required ``prst`` attribute."""
        return self.get('prst')

    def rewrite_guides(self, guides):
        """
        Remove any ``<a:gd>`` element children of ``<a:avLst>`` and replace
        them with ones having (name, val) in *guides*.
        """
        try:
            avLst = self.avLst
        except AttributeError:
            avLst = SubElement(self, 'a:avLst')
        if hasattr(self.avLst, 'gd'):
            for gd_elm in self.avLst.gd[:]:
                avLst.remove(gd_elm)
        for name, val in guides:
            gd = SubElement(avLst, 'a:gd')
            gd.set('name', name)
            gd.set('fmla', 'val %d' % val)


class CT_Shape(BaseShapeElement):
    """<p:sp> custom element class"""
    _autoshape_sp_tmpl = (
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

    _ph_sp_tmpl = (
        '<p:sp %s>\n'
        '  <p:nvSpPr>\n'
        '    <p:cNvPr id="%s" name="%s"/>\n'
        '    <p:cNvSpPr/>\n'
        '    <p:nvPr/>\n'
        '  </p:nvSpPr>\n'
        '  <p:spPr/>\n'
        '</p:sp>' % (nsdecls('a', 'p'), '%d', '%s')
    )

    _textbox_sp_tmpl = (
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
        '</p:sp>' % (nsdecls('a', 'p'), '%d', '%s', '%d', '%d', '%d', '%d')
    )

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
        prstGeom = child(self.spPr, 'a:prstGeom')
        if prstGeom is None:
            return False
        txBox = self.nvSpPr.cNvSpPr.get('txBox')
        if txBox in ('true', '1'):
            return False
        return True

    @property
    def is_textbox(self):
        """
        True if this shape is a text box. A shape is a text box if it has a
        txBox="1" attribute on cNvSpPr.
        """
        txBox = self.nvSpPr.cNvSpPr.get('txBox')
        if txBox in ('true', '1'):
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
        xml = CT_Shape._autoshape_sp_tmpl % (id_, name, left, top,
                                             width, height, prst)
        sp = parse_xml_bytes(xml)
        objectify.deannotate(sp, cleanup_namespaces=True)
        return sp

    @staticmethod
    def new_placeholder_sp(id_, name, ph_type, orient, sz, idx):
        """
        Return a new ``<p:sp>`` element tree configured as a placeholder
        shape.
        """
        xml = CT_Shape._ph_sp_tmpl % (id_, name)
        sp = parse_xml_bytes(xml)

        # placeholder shapes get a "no group" lock
        SubElement(sp.nvSpPr.cNvSpPr, 'a:spLocks')
        sp.nvSpPr.cNvSpPr[qn('a:spLocks')].set('noGrp', '1')

        # placeholder (ph) element attributes values vary by type
        ph = SubElement(sp.nvSpPr.nvPr, 'p:ph')
        if ph_type != ST_PlaceholderType.OBJ:
            ph.set('type', ph_type)
        if orient != ST_Direction.HORZ:
            ph.set('orient', orient)
        if sz != ST_PlaceholderSize.FULL:
            ph.set('sz', sz)
        if idx != 0:
            ph.set('idx', str(idx))

        placeholder_types_that_have_a_text_frame = (
            ST_PlaceholderType.TITLE, ST_PlaceholderType.CTR_TITLE,
            ST_PlaceholderType.SUB_TITLE, ST_PlaceholderType.BODY,
            ST_PlaceholderType.OBJ
        )

        if ph_type in placeholder_types_that_have_a_text_frame:
            sp.append(CT_TextBody.new_txBody())

        objectify.deannotate(sp, cleanup_namespaces=True)
        return sp

    @staticmethod
    def new_textbox_sp(id_, name, left, top, width, height):
        """
        Return a new ``<p:sp>`` element tree configured as a base textbox
        shape.
        """
        xml = CT_Shape._textbox_sp_tmpl % (id_, name, left, top, width, height)
        sp = parse_xml_bytes(xml)
        objectify.deannotate(sp, cleanup_namespaces=True)
        return sp

    @property
    def prst(self):
        """
        Value of ``prst`` attribute of ``<a:prstGeom>`` element or |None| if
        not present.
        """
        prstGeom = child(self.spPr, 'a:prstGeom')
        if prstGeom is None:
            return None
        return prstGeom.get('prst')

    @property
    def prstGeom(self):
        """
        Reference to ``<a:prstGeom>`` child element or |None| if this shape
        doesn't have one, for example, if it's a placeholder shape.
        """
        return child(self.spPr, 'a:prstGeom')

    @property
    def spPr(self):
        """
        Required ``<p:spPr>`` child element containing shape properties
        """
        return self.find(qn('p:spPr'))
