# encoding: utf-8

"""
Initialize lxml parser.
"""

from __future__ import absolute_import

from lxml import etree, objectify

from pptx.oxml.ns import _NamespacePrefixedTag, nsmap, qn


# oxml-specific constants --------------
XSD_TRUE = '1'


# configure objectified XML parser
fallback_lookup = objectify.ObjectifyElementClassLookup()
element_class_lookup = etree.ElementNamespaceClassLookup(fallback_lookup)
oxml_parser = etree.XMLParser(remove_blank_text=True)
oxml_parser.set_element_class_lookup(element_class_lookup)


def child(element, child_tagname):
    """
    Return direct child of *element* having *child_tagname* or |None|
    if no such child element is present.
    """
    xpath = './%s' % child_tagname
    matching_children = element.xpath(xpath, namespaces=nsmap)
    return matching_children[0] if len(matching_children) else None


def get_or_add(start_elm, *path_tags):
    """
    Retrieve the element at the end of the branch starting at parent and
    traversing each of *path_tags* in order, creating any elements not found
    along the way. Not a good solution when sequence of added children is
    likely to be a concern.
    """
    parent = start_elm
    for tag in path_tags:
        child_ = child(parent, tag)
        if child_ is None:
            child_ = SubElement(parent, tag)
        parent = child_
    return child_


def Element(tag):
    namespace_prefixed_tag = _NamespacePrefixedTag(tag)
    tag_name = namespace_prefixed_tag.clark_name
    tag_nsmap = namespace_prefixed_tag.namespace_map
    return oxml_parser.makeelement(tag_name, nsmap=tag_nsmap)


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
    objectify.deannotate(elm, xsi=False, cleanup_namespaces=False)
    xml = etree.tostring(
        elm, encoding=encoding, pretty_print=pretty_print,
        standalone=standalone
    )
    return xml


def SubElement(parent, tag):
    namespace_prefixed_tag = _NamespacePrefixedTag(tag)
    tag_name = namespace_prefixed_tag.clark_name
    tag_nsmap = namespace_prefixed_tag.namespace_map
    return objectify.SubElement(parent, tag_name, nsmap=tag_nsmap)


def sub_elm(parent, tag, **extra):
    return objectify.SubElement(parent, qn(tag), **extra)
