#!/usr/bin/env python
# -*- coding: utf-8 -*-

# parse_xsd.py
#

"""
Experimental code to parse XML Schema files for XML object binding classes
"""

import os
import sys

from lxml import etree

thisdir, thisfilename = os.path.split(__file__)

nsmap = {
    'xsd': 'http://www.w3.org/2001/XMLSchema',
    'p':   'http://purl.oclc.org/ooxml/presentationml/main',
    'a':   'http://purl.oclc.org/ooxml/drawingml/main',
    'r':   'http://purl.oclc.org/ooxml/officeDocument/relationships',
    's':   'http://purl.oclc.org/ooxml/officeDocument/sharedTypes',
}

reverse_nsmap = {uri: prefix for prefix, uri in nsmap.items()}


def pfxdtag(tag):
    """
    Return short-form prefixed tag from fully qualified (Clark notation)
    tagname.
    """
    uri, tagroot = tag[1:].split('}')
    prefix = reverse_nsmap[uri]
    return '%s:%s' % (prefix, tagroot)


def qtag(tag):
    """
    Return fully qualified (Clark notation) tagname corresponding to
    short-form prefixed tagname *tag*.
    """
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{%s}%s' % (uri, tagroot)


# ============================================================================
# XSD Classes
# ============================================================================

class XsdTree(object):
    """
    Wrapper for ElementTree of a .xsd file
    """
    def __init__(self, xsd_tree, nsprefix):
        super(XsdTree, self).__init__()
        self._tree = xsd_tree
        self.nsprefix = nsprefix

    def xpath(self, xpath):
        return self._tree.xpath(xpath, namespaces=nsmap)


class TypeGraph(object):
    """
    Container for graph of XSD types
    """
    def __init__(self, xsd_trees):
        super(TypeGraph, self).__init__()
        self.__xsd_trees = [XsdTree(tree, nspfx) for tree, nspfx in xsd_trees]
        self.types = {}
        self.elements = {}

    def add_element(self, element):
        """Add *element* to elements dictionary, keyed by name"""
        self.elements[element.tag] = element

    def get_complexType(self, typename):
        """Return complex type element with name *typename*"""
        if typename.startswith('a:'):
            typename = typename[2:]
        xpath = "./xsd:complexType[@name='%s']" % typename
        for xsd in self.__xsd_trees:
            elms = xsd.xpath(xpath)
            if len(elms):
                return elms[0], xsd.nsprefix
        raise KeyError("no complexType named '%s' found" % typename)

    def getdef(self, defname, tag='*'):
        """Return definition element with name *defname*"""
        if defname.startswith('a:'):
            defname = defname[2:]
        for xsd in self.__xsd_trees:
            xpath = "./%s[@name='%s']" % (tag, defname)
            elements = xsd.xpath(xpath)
            if elements:
                return elements[0]
        raise KeyError("no definition named '%s' found" % defname)

    def load_schema(self, xsd_root, nsprefix):
        for elm in xsd_root:
            tag = pfxdtag(elm.tag)
            if tag in ('xsd:import', 'xsd:simpleType', 'xsd:group',
                       'xsd:attributeGroup'):
                continue
            elif tag == 'xsd:complexType':
                ct = ComplexType(self, elm, nsprefix)
                self.types[ct.name] = ct
            elif tag == 'xsd:element':
                # if <xsd:element> has a 'ref' attribute, it is a reference
                # to a global xsd:element
                ref = elm.get('ref')
                if ref is not None:
                    elm = self.getdef(ref, tag='xsd:element')
                self.add_element(Element(elm, nsprefix))
            else:
                raise TypeError("don't know how to process %s" % tag)

    def load_subgraph(self, typename, visited=None):
        complexType_elm, nsprefix = self.get_complexType(typename)
        # take care to avoid infinite recursion by only visting each type once
        visited = [] if visited is None else visited
        if typename in visited:
            return
        visited.append(typename)
        # add type to graph
        complexType = ComplexType(self, complexType_elm, nsprefix)
        self.types[complexType.name] = complexType
        # recurse on subgraph of type
        for element in complexType.elements:
            if element.typename in ('xsd:string'):
                continue
            self.load_subgraph(element.typename, visited)


class ComplexType(object):
    """
    <xsd:complexType> object
    """
    def __init__(self, type_graph, complexType_elm, nsprefix):
        super(ComplexType, self).__init__()
        self.type_graph = type_graph
        self.nsprefix = nsprefix
        self.name = complexType_elm.get('name')
        self.elements = []
        self.attributes = []

        for child_elm in complexType_elm:
            tag = pfxdtag(child_elm.tag)
            if tag == 'xsd:sequence':
                self.__process_sequence(child_elm)
            elif tag == 'xsd:choice':
                self.__expand_choice(child_elm)
            elif tag == 'xsd:group':
                ref = child_elm.get('ref')
                if ref is not None:  # reference to separate group
                    child_elm = self.type_graph.getdef(ref)
                self.__expand_group(child_elm)
            elif tag == 'xsd:attribute':
                self.attributes.append(Attribute(child_elm))
            elif tag == 'xsd:attributeGroup':
                # e.g. ``<xsd:attributeGroup ref="AG_Locking"/>``
                ref = child_elm.get('ref')
                if ref is not None:  # reference to separate group
                    child_elm = self.type_graph.getdef(ref)
                self.__expand_attributeGroup(child_elm)
            else:
                xml = etree.tostring(child_elm)
                raise TypeError("don't know how to process %s" % xml)

    def __str__(self):
        s = '%s' % self.name
        for element in self.elements:
            s += '\n  %s' % element
        for attribute in self.attributes:
            s += '\n  %s' % attribute
        return s

    def add_element(self, element):
        """
        Add an element to this ComplexType and also append it to element dict
        of parent type graph.
        """
        self.elements.append(element)
        self.type_graph.add_element(element)

    @property
    def element_def(self):
        """
        blipFill = ElementDef('p:blipFill', 'CT_BlipFillProperties')
        blipFill.add_child('a:blip', cardinality='?')
        blipFill.add_attributes('dpi', 'rotWithShape')
        """
        s = ("%s = ElementDef('%s', '%s')\n" %
             (self.name, self.name, self.name))
        for element in self.elements:
            s += ("%s.add_child('%s', cardinality='%s')\n" %
                  (self.name, element.name, element.cardinality))
        # for attribute in self.attributes:
        #     s += '\n  %s' % attribute
        return s

    @property
    def max_tagname_len(self):
        """Return length of longest child element tagname"""
        return max([len(e.name) for e in self.elements])

    @property
    def optional_attributes(self):
        """Return list of optional attributes for this type"""
        return [a for a in self.attributes if not a.is_required]

    @property
    def required_attributes(self):
        """Return list of required attributes for this type"""
        return [a for a in self.attributes if a.is_required]

    def __expand_attributeGroup(self, attrgrp_elm):
        for child_elm in attrgrp_elm:
            tag = pfxdtag(child_elm.tag)
            if tag == 'xsd:attribute':
                self.attributes.append(Attribute(child_elm))
            else:
                tmpl = "don't know how to process %s in %s"
                raise TypeError(tmpl % (tag, 'xsd:attributeGroup'))

    def __expand_choice(self, choice_elm):
        for child_elm in choice_elm:
            tag = pfxdtag(child_elm.tag)
            if tag == 'xsd:element':
                element = Element(child_elm, self.nsprefix)
                element.minOccurs = '0'
                self.add_element(element)
            else:
                tmpl = "don't know how to process %s in %s"
                raise TypeError(tmpl % (tag, 'xsd:attributeGroup'))

    def __expand_group(self, group_elm):
        for child_elm in group_elm:
            tag = pfxdtag(child_elm.tag)
            if tag == 'xsd:sequence':
                self.__process_sequence(child_elm)
            elif tag == 'xsd:choice':
                self.__expand_choice(child_elm)
            else:
                print etree.tostring(child_elm)
                tmpl = "don't know how to process %s in %s"
                raise TypeError(tmpl % (tag, 'xsd:group'))

    def __process_sequence(self, seq_elm):
        for child_elm in seq_elm:
            tag = pfxdtag(child_elm.tag)
            if tag == 'xsd:element':
                ref = child_elm.get('ref')
                if ref is not None:  # reference to global xsd:element
                    child_elm = self.type_graph.getdef(ref, tag='xsd:element')
                self.add_element(Element(child_elm, self.nsprefix))
            elif tag == 'xsd:group':
                ref = child_elm.get('ref')
                if ref is not None:  # reference to separate group
                    child_elm = self.type_graph.getdef(ref)
                self.__expand_group(child_elm)
            elif tag == 'xsd:choice':
                self.__expand_choice(child_elm)
            # ignore xsd:any elements, just placeholders for schema expansion
            elif tag in ('xsd:any'):
                continue
            else:
                xml = etree.tostring(child_elm)
                raise TypeError("don't know how to process %s" % xml)


class Attribute(object):
    """
    <xsd:attribute> object
    """
    def __init__(self, attribute_elm):
        super(Attribute, self).__init__()
        self.name = attribute_elm.get('name')
        self.type = attribute_elm.get('type')
        self.use = attribute_elm.get('use', 'optional')
        self.default = attribute_elm.get('default')
        self.ref = attribute_elm.get('ref')
        self.form = attribute_elm.get('form')
        if self.form is not None:
            raise TypeError('found xsd:attribute with form="%s"' % self.form)
        # don't care about details other than name, ref gives that
        if self.name is None:
            self.name = self.ref

    def __str__(self):
        cardinality = '?' if self.use == 'optional' else '1'
        default = (' default="%s"' % self.default) if self.default else ''
        tmpl = 'xsd:attribute %s %s%s'
        return tmpl % (cardinality, self.name, default)

    @property
    def is_required(self):
        return self.use == 'required'


class Element(object):
    """
    <xsd:element> object
    """
    def __init__(self, element_elm, nsprefix):
        super(Element, self).__init__()
        self.nsprefix = nsprefix
        self.name = element_elm.get('name')
        self.tag = '%s:%s' % (self.nsprefix, self.name)
        self.typename = element_elm.get('type')
        if self.typename.startswith('a:'):
            self.typename = self.typename[2:]
        self.default = element_elm.get('default')
        self.form = element_elm.get('form')
        self.minOccurs = element_elm.get('minOccurs', '1')
        self.maxOccurs = element_elm.get('maxOccurs', '1')
        if self.form is not None:
            raise TypeError('found xsd:element with form="%s"' % self.form)

    def __str__(self):
        tmpl = 'xsd:element %s %s   %s%s'
        default = (' default="%s"' % self.default) if self.default else ''
        return tmpl % (self.cardinality, self.name, self.typename, default)

    @property
    def cardinality(self):
        if self.minOccurs == '0':
            cardinality = '?' if self.maxOccurs == '1' else '*'
        elif self.minOccurs == '1':
            cardinality = '1' if self.maxOccurs == '1' else '+'
        else:
            cardinality = '{%s:%s}' % (self.minOccurs, self.maxOccurs)
        return cardinality


def element_defs(tags):
    def _padding(type, element):
        return ' ' * (type.max_tagname_len - len(element.name))

    s = ''
    for tag in tags:
        element = tg.elements[tag]
        # no def needed for built-in types
        if element.typename in ('xsd:string'):
            continue
        try:
            type = tg.types[element.typename]
        except KeyError:
            continue

        # e.g. blipFill = ElementDef('p:blipFill', 'CT_BlipFillProperties')
        s += ("\n%s = ElementDef('%s', '%s')\n" %
              (element.name, element.tag, element.typename))

        # e.g. blipFill.add_child('a:blip', cardinality='?')
        for child in type.elements:
            s += ("%s.add_child('%s'%s, cardinality='%s')\n" %
                  (element.name, child.tag, _padding(type, child),
                   child.cardinality))

        # e.g. ext.add_attribute('x', required=True, default="0")
        for a in type.required_attributes:
            default = a.default if a.default else ''
            s += ("%s.add_attribute('%s', required=True, default='%s')\n" %
                  (element.name, a.name, default))

        # e.g. xfrm.add_attributes('rot', 'flipH', 'flipV')
        if type.optional_attributes:
            params = "', '".join([a.name for a in type.optional_attributes])
            s += "%s.add_attributes('%s')\n" % (element.name, params)

    return s


# ============================================================================
# main
# ============================================================================

pml = etree.parse(os.path.join(thisdir, 'xsd', 'pml.xsd')).getroot()
dml = etree.parse(os.path.join(thisdir, 'xsd', 'dml-main.xsd')).getroot()
sst = etree.parse(os.path.join(thisdir, 'xsd',
                               'shared-commonSimpleTypes.xsd')).getroot()

# xpath = "./xsd:complexType[@name='CT_Shape']"
# ct_shape = pml.xpath(xpath, namespaces=nsmap)[0]

tg = TypeGraph([(pml, 'p'), (dml, 'a'), (sst, 's')])

# tg.load_subgraph('CT_Shape')
# tg.load_subgraph('CT_Slide')
tg.load_schema(pml, 'p')
tg.load_schema(dml, 'a')
tg.load_schema(sst, 's')

# for typename in sorted(tg.types.keys()):
#     print typename
#     # print tg.types[typename]

# print '\n%d distinct types' % len(tg.types)

tags = sorted(tg.elements.keys())

# for tag in tags:
#     print(tag)
#
# print '\n%d distinct tags' % len(tags)

print element_defs(tags)

# tg.load_schema(pml, 'p')
# tg.load_schema(dml, 'a')

# tags = ['p:blipFill', 'p:pic', 'a:blip']
# tags = ['a:off', 'a:ext', 'p:blipFill', 'p:pic', 'a:blip']

# print element_defs(tags)

# BACKLOG: detect namespace prefix automatically from xsd root element

sys.exit()
