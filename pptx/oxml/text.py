# encoding: utf-8

"""
lxml custom element classes for text-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml import Element, element_class_lookup, nsmap, oxml_fromstring
from pptx.oxml.ns import nsdecls


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
        txBody = oxml_fromstring(xml)
        objectify.deannotate(txBody, cleanup_namespaces=True)
        return txBody


class CT_TextParagraph(objectify.ObjectifiedElement):
    """<a:p> custom element class"""
    def get_algn(self):
        """
        Paragraph horizontal alignment value, like ``TAT.CENTER``. Value of
        algn attribute on <a:pPr> child element
        """
        if not hasattr(self, 'pPr'):
            return None
        return self.pPr.get('algn')

    def set_algn(self, value):
        """
        Set value of algn attribute on <a:pPr> child element
        """
        if value is None:
            return self._clear_algn()
        if not hasattr(self, 'pPr'):
            pPr = Element('a:pPr')
            self.insert(0, pPr)
        self.pPr.set('algn', value)

    def _clear_algn(self):
        """
        Remove algn attribute from ``<a:pPr>`` if it exists and remove
        ``<a:pPr>`` element if it then has no attributes.
        """
        if not hasattr(self, 'pPr'):
            return
        if 'algn' in self.pPr.attrib:
            del self.pPr.attrib['algn']
        if len(self.pPr.attrib) == 0:
            self.remove(self.pPr)


a_namespace = element_class_lookup.get_namespace(nsmap['a'])
a_namespace['p'] = CT_TextParagraph

p_namespace = element_class_lookup.get_namespace(nsmap['p'])
p_namespace['txBody'] = CT_TextBody
