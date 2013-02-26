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

nsmap =\
    { 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    , 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    , 'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
    }

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


# ============================================================================
# utility functions
# ============================================================================

def _child(element, child_tagname):
    """
    Return direct child of *element* having *child_tagname* or :class:`None`
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
# Custom element classes
# ============================================================================

# def _required_attribute(element, name, default):
#     """
#     Add attribute with default value to element if it doesn't already exist.
#     """
#     if element.get(name) is None:
#         element.set(name, default)
# 
# def _required_child(parent, tag):
#     """
#     Add child element with *tag* to *parent* if it doesn't already exist.
#     """
#     if _child(parent, tag) is None:
#         parent.append(_Element(tag))


# class ElementBase(etree.ElementBase):
#     """Provides the base element interface for custom element classes"""
#     pass
# 
# class CT_Presentation(ElementBase):
#     """<p:presentation> custom element class"""
#     def _init(self):
#         _required_child(self, 'p:notesSz')
# 
#     # child accessors -----------------
#     notesSz            = property(lambda self: _child(self, 'p:notesSz'))
#     sldMasterIdLst     = property(lambda self: _get_child_or_append(self, 'p:sldMasterIdLst'))
#     notesMasterIdLst   = property(lambda self: _get_child_or_append(self, 'p:notesMasterIdLst'))
#     handoutMasterIdLst = property(lambda self: _get_child_or_append(self, 'p:handoutMasterIdLst'))
#     sldIdLst           = property(lambda self: _get_child_or_append(self, 'p:sldIdLst'))
#     sldSz              = property(lambda self: _get_child_or_append(self, 'p:sldSz'))
#     smartTags          = property(lambda self: _get_child_or_append(self, 'p:smartTags'))
#     embeddedFontLst    = property(lambda self: _get_child_or_append(self, 'p:embeddedFontLst'))
#     custShowLst        = property(lambda self: _get_child_or_append(self, 'p:custShowLst'))
#     photoAlbum         = property(lambda self: _get_child_or_append(self, 'p:photoAlbum'))
#     custDataLst        = property(lambda self: _get_child_or_append(self, 'p:custDataLst'))
#     kinsoku            = property(lambda self: _get_child_or_append(self, 'p:kinsoku'))
#     defaultTextStyle   = property(lambda self: _get_child_or_append(self, 'p:defaultTextStyle'))
#     modifyVerifier     = property(lambda self: _get_child_or_append(self, 'p:modifyVerifier'))
#     extLst             = property(lambda self: _get_child_or_append(self, 'p:extLst'))
# 
#     # attribute accessors -------------
#     serverZoom               = property(lambda self: self.get('serverZoom'),
#                                         lambda self, value: self.set('serverZoom', value))
#     firstSlideNum            = property(lambda self: self.get('firstSlideNum'),
#                                         lambda self, value: self.set('firstSlideNum', value))
#     showSpecialPlsOnTitleSld = property(lambda self: self.get('showSpecialPlsOnTitleSld'),
#                                         lambda self, value: self.set('showSpecialPlsOnTitleSld', value))
#     rtl                      = property(lambda self: self.get('rtl'),
#                                         lambda self, value: self.set('rtl', value))
#     removePersonalInfoOnSave = property(lambda self: self.get('removePersonalInfoOnSave'),
#                                         lambda self, value: self.set('removePersonalInfoOnSave', value))
#     compatMode               = property(lambda self: self.get('compatMode'),
#                                         lambda self, value: self.set('compatMode', value))
#     strictFirstAndLastChars  = property(lambda self: self.get('strictFirstAndLastChars'),
#                                         lambda self, value: self.set('strictFirstAndLastChars', value))
#     embedTrueTypeFonts       = property(lambda self: self.get('embedTrueTypeFonts'),
#                                         lambda self, value: self.set('embedTrueTypeFonts', value))
#     saveSubsetFonts          = property(lambda self: self.get('saveSubsetFonts'),
#                                         lambda self, value: self.set('saveSubsetFonts', value))
#     autoCompressPictures     = property(lambda self: self.get('autoCompressPictures'),
#                                         lambda self, value: self.set('autoCompressPictures', value))
#     bookmarkIdSeed           = property(lambda self: self.get('bookmarkIdSeed'),
#                                         lambda self, value: self.set('bookmarkIdSeed', value))
#     conformance              = property(lambda self: self.get('conformance'),
#                                         lambda self, value: self.set('conformance', value))
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


