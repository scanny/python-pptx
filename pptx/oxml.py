# -*- coding: utf-8 -*-
#
# oxml.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""
Classes that directly manipulate Open XML and provide direct object-oriented
access to the XML elements. Classes are implemented as a wrapper around their
bit of the lxml graph that spans the entire Open XML package part, e.g. a
slide.
"""

from lxml import etree, objectify

nsmap = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}

etree.register_namespace('a', nsmap['a'])
etree.register_namespace('p', nsmap['p'])

oxml_parser = objectify.makeparser(remove_blank_text=True)


# ============================================================================
# API functions
# ============================================================================

def _Element(tag, nsmap=None):
    return objectify.Element(qn(tag), nsmap=nsmap)


def _SubElement(parent, tag, nsmap=None):
    return objectify.SubElement(parent, qn(tag), nsmap=nsmap)


def new(tag, **extra):
    return objectify.Element(qn(tag), **extra)


def nsdecls(*prefixes):
    return ' '.join(['xmlns:%s="%s"' % (pfx, nsmap[pfx]) for pfx in prefixes])


def oxml_fromstring(text):
    """``etree.fromstring()`` replacement that uses oxml parser"""
    return objectify.fromstring(text, oxml_parser)


def oxml_parse(source):
    """``etree.parse()`` replacement that uses oxml parser"""
    return objectify.parse(source, oxml_parser)


def oxml_tostring(elm, encoding=None, pretty_print=False, standalone=None):
    # if xsi parameter is not set to False, PowerPoint won't load without a
    # repair step; deannotate removes some original xsi:type tags in core.xml
    # if this parameter is left out (or set to True)
    objectify.deannotate(elm, xsi=False, cleanup_namespaces=True)
    return etree.tostring(elm, encoding=encoding, pretty_print=pretty_print,
                          standalone=standalone)


def qn(tag):
    """
    Stands for "qualified name", a utility function to turn a namespace
    prefixed tag name into a Clark-notation qualified tag name for lxml. For
    example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
    """
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{%s}%s' % (uri, tagroot)


def sub_elm(parent, tag, **extra):
    return objectify.SubElement(parent, qn(tag), **extra)


# ============================================================================
# utility functions
# ============================================================================

def _child(element, child_tagname):
    """
    Return direct child of *element* having *child_tagname* or |None|
    if no such child element is present.
    """
    xpath = './%s' % child_tagname
    matching_children = element.xpath(xpath, namespaces=nsmap)
    return matching_children[0] if len(matching_children) else None


def _child_list(element, child_tagname):
    """
    Return list containing the direct children of *element* having
    *child_tagname*.
    """
    xpath = './%s' % child_tagname
    return element.xpath(xpath, namespaces=nsmap)


def _get_or_add(start_elm, *path_tags):
    """
    Retrieve the element at the end of the branch starting at parent and
    traversing each of *path_tags* in order, creating any elements not found
    along the way. Not a good solution when sequence of added children is
    likely to be a concern.
    """
    parent = start_elm
    for tag in path_tags:
        child = _child(parent, tag)
        if child is None:
            child = _SubElement(parent, tag, nsmap)
        parent = child
    return child


# ============================================================================
# Element constructors
# ============================================================================

def _empty_cell():
    tc = new('a:tc')
    tc.txBody = new('a:txBody')
    tc.txBody.bodyPr = new('a:bodyPr')
    tc.txBody.lstStyle = new('a:lstStyle')
    tc.txBody.p = new('a:p')
    tc.tcPr = new('a:tcPr')
    return tc


def CT_GraphicalObjectFrame(id, name, rows, cols, left, top, width, height):
    """
    Corresponds to the ``<p:graphicFrame>`` element, which represents a table
    shape.
    """
    xml = graphicFrame_tmpl % (id, name, left, top, width, height)
    graphicFrame = oxml_fromstring(xml)

    tbl = graphicFrame[qn('a:graphic')].graphicData.tbl
    rowheight = height/rows
    colwidth = width/cols

    for col in range(cols):
        # adjust width of last col to absorb any div error
        if col == cols-1:
            colwidth = width - ((cols-1) * colwidth)
        sub_elm(tbl.tblGrid, 'a:gridCol', w=str(colwidth))

    for row in range(rows):
        # adjust height of last row to absorb any div error
        if row == rows-1:
            rowheight = height - ((rows-1) * rowheight)
        tr = sub_elm(tbl, 'a:tr', h=str(rowheight))
        for col in range(cols):
            tr.append(_empty_cell())

    objectify.deannotate(graphicFrame, cleanup_namespaces=True)
    return graphicFrame


# ============================================================================
# Element templates
# ============================================================================

_uri = 'http://schemas.openxmlformats.org/drawingml/2006/table'
_guid = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

graphicFrame_tmpl = """
    <p:graphicFrame %s>
      <p:nvGraphicFramePr>
        <p:cNvPr id="%s" name="%s"/>
        <p:cNvGraphicFramePr>
          <a:graphicFrameLocks noGrp="1"/>
        </p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="%s" y="%s"/>
        <a:ext cx="%s" cy="%s"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="%s">
          <a:tbl>
            <a:tblPr firstRow="1" bandRow="1">
              <a:tableStyleId>%s</a:tableStyleId>
            </a:tblPr>
            <a:tblGrid/>
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>""" % (nsdecls('p', 'a'), '%d', '%s', '%d', '%d', '%d',
                            '%d', _uri, _guid)


# ============================================================================
# Custom element classes
# ============================================================================

# def _required_attribute(element, name, default):
#     """
#     Add attribute with default value to element if it doesn't already exist.
#     """
#     if element.get(name) is None:
#         element.set(name, default)
#
#
# def _required_child(parent, tag):
#     """
#     Add child element with *tag* to *parent* if it doesn't already exist.
#     """
#     if _child(parent, tag) is None:
#         parent.append(_Element(tag))
#
#
# class ElementBase(etree.ElementBase):
#     """Provides the base element interface for custom element classes"""
#     pass
#
#
# class CT_Presentation(ElementBase):
#     """<p:presentation> custom element class"""
#     def _init(self):
#         _required_child(self, 'p:notesSz')
#
#     # child accessors -----------------
#     notesSz = property(lambda self: _child(self, 'p:notesSz'))
#     sldSz = property(lambda self: _get_child_or_append(self, 'p:sldSz'))
#
#     # attribute accessors -------------
#     serverZoom = property(lambda self: self.get('serverZoom'),
#                           lambda self, value: self.set('serverZoom', value))
#
#
# class ElementClassLookup(etree.CustomElementClassLookup):
#     cls_map =\
#         { 'pic'          : CT_Picture
#         , 'nvPicPr'      : CT_PictureNonVisual
#         , 'ph'           : CT_Placeholder
#         , 'presentation' : CT_Presentation
#         }
#
#
#     def lookup(self, node_type, document, namespace, name):
#         if name in self.cls_map:
#             return self.cls_map[name]
#         return None

# oxml_parser.set_element_class_lookup(ElementClassLookup())
