# encoding: utf-8

"""
lxml custom element classes for shape-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from .. import parse_xml_bytes
from ..dml.fill import EG_FillProperties
from ..ns import nsdecls, _nsmap, qn
from ..shapes.shared import EG_EffectProperties, EG_Geometry
from .shared import (
    BaseShapeElement, ST_Direction, ST_PlaceholderSize, ST_PlaceholderType
)
from ..shared import (
    BaseOxmlElement, child, ChildTagnames, Element, SubElement
)
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
        return self.find(qn('p:spPr'))


class CT_ShapeProperties(BaseOxmlElement):
    """
    Custom element class for <p:spPr> element. Shared by ``<p:sp>``,
    ``<p:pic>``, and ``<p:cxnSp>`` elements as well as a few more obscure
    ones.
    """

    child_tagnames = ChildTagnames.from_nested_sequence(
        'a:xfrm',
        EG_Geometry.__member_names__,
        EG_FillProperties.__member_names__,
        'a:ln',
        EG_EffectProperties.__member_names__,
        'a:scene3d', 'a:sp3d', 'a:extLst',
    )

    @property
    def cx(self):
        """
        Shape width, or None if not present.
        """
        cx_str_lst = self.xpath('./a:xfrm/a:ext/@cx', namespaces=_nsmap)
        if not cx_str_lst:
            return None
        return int(cx_str_lst[0])

    @property
    def cy(self):
        """
        Shape height, or None if not present.
        """
        cy_str_lst = self.xpath('./a:xfrm/a:ext/@cy', namespaces=_nsmap)
        if not cy_str_lst:
            return None
        return int(cy_str_lst[0])

    @property
    def eg_fill_properties(self):
        """
        Return the child representing the EG_FillProperties element group
        member in this element, or |None| if no such child is present.
        """
        return self.first_child_found_in(
            *EG_FillProperties.__member_names__
        )

    def get_or_add_xfrm(self):
        """
        Return the <a:xfrm> child element, newly added if not already
        present.
        """
        xfrm = self.xfrm
        if xfrm is None:
            xfrm = self._add_xfrm()
        return xfrm

    def get_or_change_to_noFill(self):
        """
        Return the <a:noFill> child element, replacing any other fill
        element if found, e.g. a <a:gradFill> element.
        """
        if self.noFill is not None:
            return self.noFill
        self.remove_eg_fill_properties()
        return self._add_noFill()

    def get_or_change_to_solidFill(self):
        """
        Return the <a:solidFill> child element, replacing any other fill
        element if found, e.g. a <a:gradFill> element.
        """
        if self.solidFill is not None:
            return self.solidFill
        self.remove_eg_fill_properties()
        return self._add_solidFill()

    @property
    def noFill(self):
        """
        The <a:noFill> child element, or None if not present.
        """
        return self.find(qn('a:noFill'))

    def remove_eg_fill_properties(self):
        """
        Remove the fill child element, e.g ``<a:solidFill>`` if present.
        """
        self.remove_if_present(*EG_FillProperties.__member_names__)

    @property
    def solidFill(self):
        """
        The <a:solidFill> child element, or None if not present.
        """
        return self.find(qn('a:solidFill'))

    @property
    def x(self):
        """
        The integer value of `./xfrm/off/@x` attribute, or None if not
        present.
        """
        x_str_lst = self.xpath('./a:xfrm/a:off/@x', namespaces=_nsmap)
        if not x_str_lst:
            return None
        return int(x_str_lst[0])

    @property
    def xfrm(self):
        """
        The <a:xfrm> child element, or None if not present.
        """
        return self.find(qn('a:xfrm'))

    @property
    def y(self):
        """
        The top of the shape, or None if not present.
        """
        y_str_lst = self.xpath('./a:xfrm/a:off/@y', namespaces=_nsmap)
        if not y_str_lst:
            return None
        return int(y_str_lst[0])

    def _add_noFill(self):
        """
        Return a newly added <a:noFill> child element, assuming no other fill
        EG_FillProperties element is present.
        """
        noFill = Element('a:noFill')
        successor_tagnames = self.child_tagnames_after('a:noFill')
        self.insert_element_before(noFill, *successor_tagnames)
        return noFill

    def _add_solidFill(self):
        """
        Return a newly added <a:solidFill> child element.
        """
        solidFill = Element('a:solidFill')
        successor_tagnames = self.child_tagnames_after('a:solidFill')
        self.insert_element_before(solidFill, *successor_tagnames)
        return solidFill

    def _add_xfrm(self):
        """
        Return a newly added <a:xfrm> child element.
        """
        xfrm = Element('a:xfrm')
        successor_tagnames = self.child_tagnames_after('a:xfrm')
        self.insert_element_before(xfrm, *successor_tagnames)
        return xfrm
