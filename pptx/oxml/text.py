# encoding: utf-8

"""
lxml custom element classes for text-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml import parse_xml_bytes
from pptx.oxml.core import Element, SubElement
from pptx.oxml.ns import nsdecls, qn


class CT_TextBody(objectify.ObjectifiedElement):
    """<p:txBody> custom element class"""
    _txBody_tmpl = (
        '<p:txBody %s>\n'
        '  <a:bodyPr/>\n'
        '  <a:lstStyle/>\n'
        '  <a:p/>\n'
        '</p:txBody>\n' % (nsdecls('a', 'p'))
    )

    @staticmethod
    def new_txBody():
        """Return a new ``<p:txBody>`` element tree"""
        xml = CT_TextBody._txBody_tmpl
        txBody = parse_xml_bytes(xml)
        objectify.deannotate(txBody, cleanup_namespaces=True)
        return txBody


class CT_TextParagraph(objectify.ObjectifiedElement):
    """
    <a:p> custom element class
    """
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

    def get_or_add_pPr(self):
        """
        Return the <a:pPr> child element of this <a:p> element, a newly added
        one if one is not present.
        """
        if not hasattr(self, 'pPr'):
            pPr = Element('a:pPr')
            self.insert(0, pPr)
        return self.pPr


class CT_TextParagraphProperties(objectify.ObjectifiedElement):
    """
    <a:pPr> custom element class
    """
    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in ('algn',):
            self._set_attr(name, value)
        else:
            super(CT_TextParagraphProperties, self).__setattr__(name, value)

    @property
    def algn(self):
        """
        Paragraph horizontal alignment value, like ``TAT.CENTER``. Value of
        'algn' attribute on <a:pPr> child element. None if no 'algn'
        attribute is present.
        """
        return self.get('algn')

    def _set_attr(self, name, value):
        """
        Set attribute of this element having *name* to *value*.
        """
        if value is None and name in self.attrib:
            del self.attrib[name]
        else:
            self.set(name, value)
