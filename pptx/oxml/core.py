# encoding: utf-8

"""
General purpose functions that raise the abstraction level of interacting with
lxml.objectify elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml import oxml_parser
from pptx.oxml.ns import NamespacePrefixedTag


def child(element, child_tag_str):
    """
    Return the first direct child of *element* having tag matching
    *child_tag_str* or |None| if no such child element is present.
    """
    nsptag = NamespacePrefixedTag(child_tag_str)
    xpath = './%s' % child_tag_str
    matching_children = element.xpath(xpath, namespaces=nsptag.nsmap)
    return matching_children[0] if len(matching_children) else None


def Element(tag):
    """
    Return 'loose' lxml element having the tag specified by *tag*. *tag*
    must contain the standard namespace prefix, e.g. 'a:tbl'. The namespace
    URI is automatically looked up from the prefix and the resulting element
    will be an instance of the custom element class for this tag name if one
    is defined.
    """
    nsptag = NamespacePrefixedTag(tag)
    return oxml_parser.makeelement(nsptag.clark_name, nsmap=nsptag.nsmap)


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


def SubElement(parent, tag, **extra):
    """
    Return an lxml element having *tag*, newly added as a direct child of
    *parent*. The new element is appended to the sequence of children, so
    this method is not suitable if the child element must be inserted at a
    different position in the sequence. The class of the returned element is
    the custom element class for its tag, if one is defined.
    """
    nsptag = NamespacePrefixedTag(tag)
    return objectify.SubElement(
        parent, nsptag.clark_name, nsmap=nsptag.nsmap, **extra
    )
