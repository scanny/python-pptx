# encoding: utf-8

"""
lxml custom element classes for text-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml import parse_xml_bytes
from pptx.oxml.core import Element, SubElement
from pptx.oxml.ns import nsdecls, nsmap, qn


class CT_Hyperlink(objectify.ObjectifiedElement):
    """
    Custom element class for <a:hlinkClick> elements.
    """
    @property
    def rId(self):
        return self.get(qn('r:id'))


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

    @property
    def bodyPr(self):
        return self[qn('a:bodyPr')]


class Schema(object):
    def __init__(self):
        self._attrs = {}

    def add_attr(self, attr):
        self._attrs[attr.prop_name] = attr

    @property
    def attrs(self):
        return self._attrs

    def is_attr_prop(self, prop_name):
        return prop_name in self._attrs.keys()


class Attribute(object):
    def __init__(self, prop_name, type_cls, attr_name=None):
        self._prop_name = prop_name
        self._type_cls = type_cls
        self._attr_name = attr_name if attr_name is not None else prop_name

    @property
    def prop_name(self):
        return self._prop_name

    def get(self, elm):
        if self._attr_name not in elm.attrib:
            return None
        xml_val = elm.get(self._attr_name)
        return self._type_cls.from_xml(xml_val)

    def set(self, elm, value):
        if value is None and self._attr_name in elm.attrib:
            del elm.attrib[self._attr_name]
        else:
            xml_val = self._type_cls.to_xml(value)
            elm.set(self._attr_name, xml_val)


class ST_Coordinate32(object):
    @classmethod
    def from_xml(cls, value):
        return int(value)

    @classmethod
    def to_xml(cls, value):
        if not isinstance(value, int):
            raise ValueError('int value required')
        return str(value)


class OxmlElement(objectify.ObjectifiedElement):

    def __getattr__(self, name):
        if self.schema.is_attr_prop(name):
            attr = self.schema.attrs[name]
            return attr.get(self)
        else:
            return super(OxmlElement, self).__getattr__(name)

    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if self.schema.is_attr_prop(name):
            attr = self.schema.attrs[name]
            attr.set(self, value)
        else:
            super(OxmlElement, self).__setattr__(name, value)


class CT_TextBodyProperties(OxmlElement):
    """
    <a:bodyPr> custom element class
    """
    schema = Schema()
    schema.add_attr(Attribute('lIns', ST_Coordinate32))
    schema.add_attr(Attribute('tIns', ST_Coordinate32))
    schema.add_attr(Attribute('rIns', ST_Coordinate32))
    schema.add_attr(Attribute('bIns', ST_Coordinate32))
    # lIns = Attribute(ST_Coordinate32)


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
        elif name == 'hlinkClick':
            return self.find(qn('a:hlinkClick'))
        else:
            return super(CT_TextCharacterProperties, self).__getattr__(name)

    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in ('b', 'i'):
            self._set_bool_attr(name, value)
        elif name == 'hlinkClick':
            self._set_hlinkClick(value)
        else:
            super(CT_TextCharacterProperties, self).__setattr__(name, value)

    def add_hlinkClick(self, rId):
        """
        Add an <a:hlinkClick> child element with r:id attribute set to *rId*.
        """
        assert self.find(qn('a:hlinkClick')) is None

        hlinkClick = Element('a:hlinkClick', nsmap('a', 'r'))
        hlinkClick.set(qn('r:id'), rId)

        # find right insertion spot, will go away once xmlchemy comes in
        if self.find(qn('a:hlinkMouseOver')):
            successor = self.find(qn('a:hlinkMouseOver'))
            successor.addprevious(hlinkClick)
        elif self.find(qn('a:rtl')):
            successor = self.find(qn('a:rtl'))
            successor.addprevious(hlinkClick)
        elif self.find(qn('a:extLst')):
            successor = self.find(qn('a:extLst'))
            successor.addprevious(hlinkClick)
        else:
            self.append(hlinkClick)

        return hlinkClick

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

    def _remove_if_present(self, tagnames):
        for tagname in tagnames:
            element = self.find(qn(tagname))
            if element is not None:
                self.remove(element)

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

    def _set_hlinkClick(self, value):
        """
        For *value* is None, remove the ``<a:hlinkClick>`` child. For all
        other values, raise |ValueError|.
        """
        if value is not None:
            tmpl = "only None can be assigned to optional element, got '%s'"
            raise ValueError(tmpl % value)
        # value is None ----------------
        hlinkClick = self.find(qn('a:hlinkClick'))
        if hlinkClick is not None:
            self.remove(hlinkClick)


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
