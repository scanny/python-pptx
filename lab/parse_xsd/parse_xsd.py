#!/usr/bin/env python
# -*- coding: utf-8 -*-

# test.py
#
# Assurance testing during the course of development

import os
import sys

from lxml import etree

import logging

thisdir, thisfilename = os.path.split(__file__)

log = logging.getLogger(thisfilename)
log.setLevel(logging.DEBUG)
# log.setLevel(logging.INFO)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)
log.addHandler(ch)

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
# parse-y types of things
# ============================================================================

# parse xsd file ------------------------------------------

# filename = 'pml.xsd'
# # filename = 'dml-main.xsd'
# path = os.path.join(thisdir, 'xsd', filename)
# xsd = etree.parse(path).getroot()

pml = etree.parse(os.path.join(thisdir, 'xsd', 'pml.xsd'     )).getroot()
dml = etree.parse(os.path.join(thisdir, 'xsd', 'dml-main.xsd')).getroot()


# ============================================================================
# XSD Classes
# ============================================================================

class ComplexType(object):
    """
    <xsd:complexType> object
    """
    def __init__(self, complexType_elm):
        super(ComplexType, self).__init__()
        self.name = complexType_elm.get('name')
    
    def __str__(self):
        return '%s' % self.name
    

# walk definition tree for CT_Shape -----------------------

def process_attribute(attr_elm):
    if len(attr_elm):
        msg = "don't know how to process xsd:attribute with children"
        raise TypeError(msg)
    tag = pfxdtag(attr_elm.tag)
    name = attr_elm.get('name')
    type = attr_elm.get('type')
    use = attr_elm.get('use', 'optional')
    default = attr_elm.get('default')
    print '  %s %s %s' % (tag, name, type)

def process_attributeGroup(attr_grp_elm):
    print '  xsd:attributeGroup'
    ref = attr_grp_elm.get('ref')
    xpath = "./xsd:attributeGroup[@name='%s']" % ref
    try:
        ag = pml.xpath(xpath, namespaces=nsmap)[0]
    except IndexError:
        ag = dml.xpath(xpath, namespaces=nsmap)[0]
    for child_elm in ag:
        tag = pfxdtag(child_elm.tag)
        # print '    %s' % tag
        if tag == 'xsd:attribute':
            process_attribute(child_elm)
        else:
            raise TypeError("don't know how to process %s" % tag)
    print '  /xsd:attributeGroup'

def process_choice(choice_elm):
    print '  xsd:choice'
    for child_elm in choice_elm:
        tag = pfxdtag(child_elm.tag)
        if tag == 'xsd:element':
            process_element(child_elm, indent=2)
        else:
            raise TypeError("don't know how to process %s" % tag)
    print '  /xsd:choice'

def process_element(elm_elm, indent=0):
    indentation = ' ' * indent
    tag = pfxdtag(elm_elm.tag)
    name = elm_elm.get('name')
    type = elm_elm.get('type')
    print '%s    %s %-10s %s' % (indentation, tag, name, type)

def process_group(grp_elm):
    tag = pfxdtag(grp_elm.tag)
    ref = grp_elm.get('ref')
    minOccurs = grp_elm.get('minOccurs')
    maxOccurs = grp_elm.get('maxOccurs')
    tmpl = '    %s %-10s minOccurs="%s" maxOccurs="%s"'
    print tmpl % (tag, ref, minOccurs, maxOccurs)

def process_sequence(seq_elm):
    print '  xsd:sequence'
    for child_elm in seq_elm:
        tag = pfxdtag(child_elm.tag)
        # print '    %s' % tag
        if tag == 'xsd:element':
            process_element(child_elm)
        elif tag == 'xsd:group':
            process_group(child_elm)
        else:
            raise TypeError("don't know how to process %s" % tag)
    print '  /xsd:sequence'

def walk_typedefs(complexType_elm, visited=None):
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


xpath = "./xsd:complexType[@name='CT_Shape']"
ct_shape = pml.xpath(xpath, namespaces=nsmap)[0]

visited = []
walk_typedefs(ct_shape, visited)

print '\n%d distinct types:' % len(visited)
for type_name in sorted(visited):
    print '  %s' % type_name

sys.exit()


name = ct_shape.get('name')
print '%s' % name

for child_elm in ct_shape:
    tag = pfxdtag(child_elm.tag)
    if tag == 'xsd:sequence':
        process_sequence(child_elm)
    elif tag == 'xsd:attribute':
        process_attribute(child_elm)
    else:
        raise TypeError("don't know how to process %s" % tag)





# ============================================================================
# code templates
# ============================================================================

# log.debug("xsd_filename -> '%s'", xsd_filename)

# # process all child elements ------------------------------
# for top_level_elm in xsd:
#     tag = top_level_elm.tag
#     print pfxdtag(tag)

# # process all type definitions ----------------------------
# for top_level_elm in xsd:
#     tag = pfxdtag(top_level_elm.tag)
#     if tag in ('xsd:complexType', 'xsd:simpleType'):
#         name = top_level_elm.get('name')
#         print '%-16s %s' % (tag, name)

# # find particular type definition -------------------------
# xpath = "./xsd:complexType[@name='CT_Shape']"
# ct_shape = xsd.xpath(xpath, namespaces=nsmap)[0]
# print ct_shape

# # detect all used child element types for complexType -----
# complexType_elms = [elm for elm in xsd if elm.tag == qtag('xsd:complexType')]
# other_types = []
# 
# for complexType in complexType_elms:
#     # print '%s %s' % (pfxdtag(complexType.tag), complexType.get('name'))
#     for child_elm in complexType:
#         tag = pfxdtag(child_elm.tag)
#         if tag in ('xsd:sequence', 'xsd:attribute'):
#             continue
#         if tag not in other_types:
#             other_types.append(tag)
#         # print '    %s' % tag
# 
# print other_types

