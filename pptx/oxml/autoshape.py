# encoding: utf-8

"""
lxml custom element classes for shape-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml import parse_xml_bytes
from pptx.oxml.core import child, Element, SubElement
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.text import CT_TextBody
from pptx.spec import (
    PH_ORIENT_HORZ, PH_SZ_FULL, PH_TYPE_BODY, PH_TYPE_CTRTITLE, PH_TYPE_OBJ,
    PH_TYPE_SUBTITLE, PH_TYPE_TITLE
)


class CT_Point2D(objectify.ObjectifiedElement):
    """
    Custom element class for <a:off> element.
    """
    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in ('x', 'y'):
            self.set(name, str(value))
        else:
            super(CT_Point2D, self).__setattr__(name, value)

    @property
    def x(self):
        """
        Integer value of required ``x`` attribute.
        """
        x_str = self.get('x')
        return int(x_str)

    @property
    def y(self):
        """
        Integer value of required ``y`` attribute.
        """
        y_str = self.get('y')
        return int(y_str)


class CT_PositiveSize2D(objectify.ObjectifiedElement):
    """
    Custom element class for <a:ext> element.
    """
    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in ('cx', 'cy'):
            self.set(name, str(value))
        else:
            super(CT_PositiveSize2D, self).__setattr__(name, value)

    @property
    def cx(self):
        """
        Integer value of required ``cx`` attribute.
        """
        cx_str = self.get('cx')
        return int(cx_str)

    @property
    def cy(self):
        """
        Integer value of required ``cy`` attribute.
        """
        cy_str = self.get('cy')
        return int(cy_str)


class CT_PresetGeometry2D(objectify.ObjectifiedElement):
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


class CT_Shape(objectify.ObjectifiedElement):
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
        if ph_type != PH_TYPE_OBJ:
            ph.set('type', ph_type)
        if orient != PH_ORIENT_HORZ:
            ph.set('orient', orient)
        if sz != PH_SZ_FULL:
            ph.set('sz', sz)
        if idx != 0:
            ph.set('idx', str(idx))

        placeholder_types_that_have_a_text_frame = (
            PH_TYPE_TITLE, PH_TYPE_CTRTITLE, PH_TYPE_SUBTITLE, PH_TYPE_BODY,
            PH_TYPE_OBJ)

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


class CT_ShapeProperties(objectify.ObjectifiedElement):
    """
    Custom element class for <p:spPr> element.
    """
    @property
    def eg_fillproperties(self):
        """
        Return the child representing the EG_FillProperties element group
        member in this element, or |None| if no such child is present.
        """
        return self._first_child_found_in(
            'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill',
            'a:pattFill', 'a:grpFill'
        )

    def get_or_add_xfrm(self):
        """
        Return the <a:xfrm> child element, newly added if not already
        present.
        """
        xfrm = self.xfrm
        if xfrm is None:
            xfrm = Element('a:xfrm')
            self.insert(0, xfrm)
        return xfrm

    def get_or_change_to_noFill(self):
        """
        Return the <a:noFill> child element, replacing any other fill
        element if found, e.g. a <a:gradFill> element.
        """
        # return existing one if there is one
        if self.noFill is not None:
            return self.noFill
        # get rid of other fill element type if there is one
        self._remove_if_present(
            'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill',
            'a:grpFill'
        )
        # add noFill element in right sequence
        return self._add_noFill()

    def get_or_change_to_solidFill(self):
        """
        Return the <a:solidFill> child element, replacing any other fill
        element if found, e.g. a <a:gradFill> element.
        """
        # return existing one if there is one
        if self.solidFill is not None:
            return self.solidFill
        # get rid of other fill element type if there is one
        self._remove_if_present(
            'a:noFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'
        )
        # add solidFill element in right sequence
        return self._add_solidFill()

    @property
    def noFill(self):
        """
        The <a:noFill> child element, or None if not present.
        """
        return self.find(qn('a:noFill'))

    @property
    def solidFill(self):
        """
        The <a:solidFill> child element, or None if not present.
        """
        return self.find(qn('a:solidFill'))

    @property
    def xfrm(self):
        """
        The <a:xfrm> child element, or None if not present.
        """
        return self.find(qn('a:xfrm'))

    def _add_noFill(self):
        """
        Return a newly added <a:noFill> child element, assuming no other fill
        EG_FillProperties element is present.
        """
        noFill = Element('a:noFill')

        successor = self._first_successor_in(
            'a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d',
            'a:extLst'
        )
        if successor is not None:
            successor.addprevious(noFill)
        else:
            self.append(noFill)

        return noFill

    def _add_solidFill(self):
        """
        Return a newly added <a:solidFill> child element.
        """
        solidFill = Element('a:solidFill')

        successor = self._first_successor_in(
            'a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d',
            'a:extLst'
        )
        if successor is not None:
            successor.addprevious(solidFill)
        else:
            self.append(solidFill)

        return solidFill

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

    def _first_successor_in(self, *successor_tagnames):
        """
        Return the first child with tag in *successor_tagnames*, or None if
        not found.
        """
        for tagname in successor_tagnames:
            successor = self.find(qn(tagname))
            if successor is not None:
                return successor
        return None

    def _remove_if_present(self, *tagnames):
        for tagname in tagnames:
            element = self.find(qn(tagname))
            if element is not None:
                self.remove(element)


class CT_Transform2D(objectify.ObjectifiedElement):
    """
    Custom element class for <a:xfrm> element.
    """
    @property
    def ext(self):
        """
        The <a:ext> child element, or None if not present.
        """
        return self.find(qn('a:ext'))

    def get_or_add_ext(self):
        """
        Return the <a:ext> child element, newly added if not already
        present.
        """
        ext = self.ext
        if ext is None:
            ext = Element('a:ext')
            ext.set('cx', '0')
            ext.set('cy', '0')
            self.append(ext)
        return ext

    def get_or_add_off(self):
        """
        Return the <a:off> child element, newly added if not already
        present.
        """
        off = self.off
        if off is None:
            off = Element('a:off')
            off.set('x', '0')
            off.set('y', '0')
            self.insert(0, off)
        return off

    @property
    def off(self):
        """
        The <a:off> child element, or None if not present.
        """
        return self.find(qn('a:off'))
