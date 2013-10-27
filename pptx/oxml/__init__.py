# encoding: utf-8

"""
Initialize lxml parser.
"""

from __future__ import absolute_import

from lxml import etree, objectify

from pptx.oxml.ns import NamespacePrefixedTag, qn


# oxml-specific constants --------------
XSD_TRUE = '1'


# configure objectified XML parser
fallback_lookup = objectify.ObjectifyElementClassLookup()
element_class_lookup = etree.ElementNamespaceClassLookup(fallback_lookup)
oxml_parser = etree.XMLParser(remove_blank_text=True)
oxml_parser.set_element_class_lookup(element_class_lookup)


def Element(tag):
    nsptag = NamespacePrefixedTag(tag)
    return oxml_parser.makeelement(nsptag.clark_name, nsmap=nsptag.nsmap)


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


def sub_elm(parent, tag, **extra):
    return objectify.SubElement(parent, qn(tag), **extra)
