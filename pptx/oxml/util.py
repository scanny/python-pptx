# encoding: utf-8

"""
Classes that directly manipulate Open XML and provide direct object-oriented
access to the XML elements. Classes are implemented as a wrapper around their
bit of the lxml graph that spans the entire Open XML package part, e.g. a
slide.
"""

from __future__ import absolute_import

from lxml import etree, objectify

from pptx.oxml import element_class_lookup, oxml_parser
from pptx.spec import nsmap


# ============================================================================
# API functions
# ============================================================================

class _NamespacePrefixedTag(str):
    """
    Value object that knows the semantics of an XML tag having a namespace
    prefix.
    """
    def __new__(cls, nstag, *args):
        return super(_NamespacePrefixedTag, cls).__new__(cls, nstag)

    def __init__(self, nstag, prefix_to_uri_map):
        self._pfx, self._local_part = nstag.split(':')
        self._ns_uri = prefix_to_uri_map[self._pfx]

    @property
    def clark_name(self):
        return '{%s}%s' % (self._ns_uri, self._local_part)

    @property
    def namespace_map(self):
        return {self._pfx: self._ns_uri}


def _Element(tag):
    namespace_prefixed_tag = _NamespacePrefixedTag(tag, nsmap)
    tag_name = namespace_prefixed_tag.clark_name
    tag_nsmap = namespace_prefixed_tag.namespace_map
    return oxml_parser.makeelement(tag_name, nsmap=tag_nsmap)


def _SubElement(parent, tag):
    namespace_prefixed_tag = _NamespacePrefixedTag(tag, nsmap)
    tag_name = namespace_prefixed_tag.clark_name
    tag_nsmap = namespace_prefixed_tag.namespace_map
    return objectify.SubElement(parent, tag_name, nsmap=tag_nsmap)


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
    objectify.deannotate(elm, xsi=False, cleanup_namespaces=False)
    xml = etree.tostring(
        elm, encoding=encoding, pretty_print=pretty_print,
        standalone=standalone
    )
    return xml


def qn(namespace_prefixed_tag):
    """
    Return a Clark-notation qualified tag name corresponding to
    *namespace_prefixed_tag*, a string like 'p:body'. 'qn' stands for
    *qualified name*. As an example, ``qn('p:cSld')`` returns
    ``'{http://schemas.../main}cSld'``.
    """
    nsptag = _NamespacePrefixedTag(namespace_prefixed_tag, nsmap)
    return nsptag.clark_name


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
            child = _SubElement(parent, tag)
        parent = child
    return child


# ============================================================================
# Custom element classes
# ============================================================================

class CT_Picture(objectify.ObjectifiedElement):
    """
    ``<p:pic>`` element, which represents a picture shape (an image placement
    on a slide).
    """
    _pic_tmpl = (
        '<p:pic %s>\n'
        '  <p:nvPicPr>\n'
        '    <p:cNvPr id="%s" name="%s" descr="%s"/>\n'
        '    <p:cNvPicPr>\n'
        '      <a:picLocks noChangeAspect="1"/>\n'
        '    </p:cNvPicPr>\n'
        '    <p:nvPr/>\n'
        '  </p:nvPicPr>\n'
        '  <p:blipFill>\n'
        '    <a:blip r:embed="%s"/>\n'
        '    <a:stretch>\n'
        '      <a:fillRect/>\n'
        '    </a:stretch>\n'
        '  </p:blipFill>\n'
        '  <p:spPr>\n'
        '    <a:xfrm>\n'
        '      <a:off x="%s" y="%s"/>\n'
        '      <a:ext cx="%s" cy="%s"/>\n'
        '    </a:xfrm>\n'
        '    <a:prstGeom prst="rect">\n'
        '      <a:avLst/>\n'
        '    </a:prstGeom>\n'
        '  </p:spPr>\n'
        '</p:pic>' % (nsdecls('a', 'p', 'r'), '%d', '%s', '%s', '%s',
                      '%d', '%d', '%d', '%d')
    )

    @staticmethod
    def new_pic(id_, name, desc, rId, left, top, width, height):
        """
        Return a new ``<p:pic>`` element tree configured with the supplied
        parameters.
        """
        xml = CT_Picture._pic_tmpl % (id_, name, desc, rId,
                                      left, top, width, height)
        pic = oxml_fromstring(xml)

        objectify.deannotate(pic, cleanup_namespaces=True)
        return pic

p_namespace = element_class_lookup.get_namespace(nsmap['p'])
p_namespace['pic'] = CT_Picture
