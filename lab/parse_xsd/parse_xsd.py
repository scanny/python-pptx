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

nsmap =\
    { 'xsd' : 'http://www.w3.org/2001/XMLSchema'
    , 'p'   : 'http://purl.oclc.org/ooxml/presentationml/main'
    , 'a'   : 'http://purl.oclc.org/ooxml/drawingml/main'
    , 'r'   : 'http://purl.oclc.org/ooxml/officeDocument/relationships'
    , 's'   : 'http://purl.oclc.org/ooxml/officeDocument/sharedTypes" elementFormDefault="qualified'
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

class TypeGraph(object):
    """
    Container for graph of XSD types
    """
    def __init__(self, xsd_trees):
        super(TypeGraph, self).__init__()
        self.__xsd_trees = xsd_trees
        self.types = []
    
    def getdef(self, defname, tag='*'):
        """Return definition element with name *defname*"""
        if defname.startswith('a:'):
            defname = defname[2:]
        for xsd in self.__xsd_trees:
            xpath = "./%s[@name='%s']" % (tag, defname)
            elements = xsd.xpath(xpath, namespaces=nsmap)
            if elements:
                return elements[0]
        raise KeyError("no definition named '%s' found" % defname)
    
    def load_schema(self, xsd_root):
        print "loading schema %s" % xsd_root
        
        for elm in xsd_root:
            tag = pfxdtag(elm.tag)
            if tag in ('xsd:import', 'xsd:simpleType', 'xsd:element',
                       'xsd:group', 'xsd:attributeGroup'):
                continue
            elif tag == 'xsd:complexType':
                ct = ComplexType(self, elm)
                self.types.append(ct)
            else:
                raise TypeError("don't know how to process %s" % tag)
        self.types.sort(key=lambda type: type.name)
    
    def walk_typedefs(self, complexType_elm, visited=None):
        type_name = complexType_elm.get('name')
        visited = [] if visited is None else visited
        if type_name in visited:
            return
        visited.append(type_name)
        
        complexType = ComplexType(complexType_elm)
        print '\n%s' % complexType
        
        for child_elm in complexType_elm:
            tag = pfxdtag(child_elm.tag)
            if tag == 'xsd:sequence':
                process_sequence(child_elm)
            elif tag == 'xsd:attribute':
                process_attribute(child_elm)
            elif tag == 'xsd:attributeGroup':
                process_attributeGroup(child_elm)
            elif tag == 'xsd:choice':
                process_choice(child_elm)
            else:
                raise TypeError("don't know how to process %s" % tag)
        
        for child_elm in complexType_elm:
            tag = pfxdtag(child_elm.tag)
            if tag != 'xsd:sequence':
                continue
            seq_elm = child_elm
            for child_elm in seq_elm:
                tag = pfxdtag(child_elm.tag)
                if tag == 'xsd:element':
                    type = child_elm.get('type')
                    if type.startswith('a:'):
                        type = type[2:]
                    xpath = "./xsd:complexType[@name='%s']" % type
                    try:
                        ct = pml.xpath(xpath, namespaces=nsmap)[0]
                    except IndexError:
                        ct = dml.xpath(xpath, namespaces=nsmap)[0]
                    walk_typedefs(ct, visited)
                elif tag == 'xsd:group':
                    pass
                else:
                    raise TypeError("don't know how to process %s" % tag)
    

class ComplexType(object):
    """
    <xsd:complexType> object
    """
    def __init__(self, type_graph, complexType_elm):
        super(ComplexType, self).__init__()
        self.type_graph = type_graph
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
                self.__expand_attributeGroup(child_elm)
            else:
                xml = etree.tostring(child_elm)
                raise TypeError("don't know how to process %s" % xml)
    
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
                element = Element(child_elm)
                element.minOccurs = '0'
                self.elements.append(element)
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
                self.elements.append(Element(child_elm))
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
    
    def __str__(self):
        s = '%s' % self.name
        for element in self.elements:
            s += '\n  %s' % element
        for attribute in self.attributes:
            s += '\n  %s' % attribute
        return s
    

class Attribute(object):
    """
    <xsd:attribute> object
    """
    def __init__(self, attribute_elm):
        super(Attribute, self).__init__()
        self.name    = attribute_elm.get('name')
        self.type    = attribute_elm.get('type')
        self.use     = attribute_elm.get('use', 'optional')
        self.default = attribute_elm.get('default')
        self.ref     = attribute_elm.get('ref')
        self.form    = attribute_elm.get('form')
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
    

class Element(object):
    """
    <xsd:element> object
    """
    def __init__(self, element_elm):
        super(Element, self).__init__()
        self.name      = element_elm.get('name')
        self.type      = element_elm.get('type')
        self.default   = element_elm.get('default')
        self.form      = element_elm.get('form')
        self.minOccurs = element_elm.get('minOccurs', '1')
        self.maxOccurs = element_elm.get('maxOccurs', '1')
        if self.form is not None:
            raise TypeError('found xsd:element with form="%s"' % self.form)
    
    def __str__(self):
        if self.minOccurs == '0':
            cardinality = '?' if self.maxOccurs == '1' else '*'
        elif self.minOccurs == '1':
            cardinality = '1' if self.maxOccurs == '1' else '+'
        else:
            cardinality = '{%s:%s}' % (self.minOccurs, self.maxOccurs)
        
        tmpl = 'xsd:element %s %s   %s%s'
        default = (' default="%s"' % self.default) if self.default else ''
        return tmpl % (cardinality, self.name, self.type, default)
    


# ============================================================================
# main
# ============================================================================

pml = etree.parse(os.path.join(thisdir, 'xsd', 'pml.xsd'     )).getroot()
dml = etree.parse(os.path.join(thisdir, 'xsd', 'dml-main.xsd')).getroot()

xpath = "./xsd:complexType[@name='CT_Shape']"
ct_shape = pml.xpath(xpath, namespaces=nsmap)[0]

tg = TypeGraph([pml, dml])

tg.load_schema(pml)
tg.load_schema(dml)
for type in tg.types:
    print type
print '\n%d distinct types' % len(tg.types)

sys.exit()

