# encoding: utf-8

"""
lxml custom element classes for text-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml import parse_xml_bytes
from pptx.oxml.core import Element, SubElement
from pptx.oxml.ns import nsdecls, qn


class CT_RegularTextRun(objectify.ObjectifiedElement):
    """
    Custom element class for <a:r> elements.
    """
    def get_or_add_rPr(self):
        """
        Return the <a:rPr> child element of this <a:r> element, newly added
        if not already present.
        """
        if not hasattr(self, 'rPr'):
            self.insert(0, Element('a:rPr'))
        return self.rPr


class CT_TextBody(objectify.ObjectifiedElement):
    """
    <p:txBody> custom element class
    """
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


class CT_TextCharacterProperties(objectify.ObjectifiedElement):
    """
    Custom element class for all of <a:rPr>, <a:defRPr>, and <a:endParaRPr>
    elements. 'rPr' is short for 'run properties', and it corresponds to the
    _Font proxy class.
    """
    def __getattr__(self, name):
        """
        Override ``__getattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in ('b', 'i'):
            return self._get_bool_attr(name)
        else:
            return super(CT_TextCharacterProperties, self).__getattr__(name)

    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in ('b', 'i'):
            self._set_bool_attr(name, value)
        else:
            super(CT_TextCharacterProperties, self).__setattr__(name, value)

    def _get_bool_attr(self, name):
        """
        True if *name* attribute is a truthy value, False if a falsey value,
        and None if no *name* attribute is present.
        """
        attr_str = self.get(name)
        if attr_str is None:
            return None
        if attr_str in ('true', '1'):
            return True
        return False

    def _set_bool_attr(self, name, value):
        """
        Set boolean attribute of this element having *name* to boolean value
        of *value*.
        """
        if value is None:
            if name in self.attrib:
                del self.attrib[name]
        elif bool(value):
            self.set(name, '1')
        else:
            self.set(name, '0')

    def get_or_change_to_solidFill(self):
        """
        Return the <a:solidFill> child element, replacing any other fill
        element if found, e.g. a <a:gradientFill> element.
        """
        # return existing one if there is one
        if self.solidFill is not None:
            return self.solidFill
        # get rid of other fill element type if there is one
        other_fill_tagnames = (
            'a:noFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'
        )
        self._remove_if_present(other_fill_tagnames)
        # add solidFill element in right sequence
        return self._add_solidFill()

    @property
    def solidFill(self):
        """
        The <a:solidFill> child element, or None if not present.
        """
        return self.find(qn('a:solidFill'))

    def _add_solidFill(self):
        """
        Return a newly added <a:solidFill> child element.
        """
        solidFill = Element('a:solidFill')
        ln = self.find(qn('a:ln'))
        if ln is not None:
            self.insert(1, solidFill)
        else:
            self.insert(0, solidFill)
        return solidFill

    def _remove_if_present(self, tagnames):
        for tagname in tagnames:
            element = self.find(qn(tagname))
            if element is not None:
                self.remove(element)


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

    def get_or_add_defRPr(self):
        """
        Return the <a:defRPr> child element of this <a:pPr> element, newly
        added if not already present.
        """
        if not hasattr(self, 'defRPr'):
            defRPr = Element('a:defRPr')
            try:
                self.extLst.addprevious(defRPr)
            except AttributeError:
                self.append(defRPr)
        return self.defRPr

    def _set_attr(self, name, value):
        """
        Set attribute of this element having *name* to *value*.
        """
        if value is None and name in self.attrib:
            del self.attrib[name]
        else:
            self.set(name, value)
